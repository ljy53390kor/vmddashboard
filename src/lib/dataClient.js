import { supabase } from './supabaseClient';
import { isVmddashboard } from './target';

// ═══════════════════════════════════════════════════════════════════════
// Vercel(기본) 경로 — 기존 Supabase 동작을 그대로 얇게 감싼다.
// ═══════════════════════════════════════════════════════════════════════
const supabaseAuth = {
  getSession: () => supabase.auth.getSession(),
  onAuthStateChange: (cb) => supabase.auth.onAuthStateChange(cb),
  signOut: () => supabase.auth.signOut(),
  signInWithPassword: (creds) => supabase.auth.signInWithPassword(creds),
  signUp: (payload) => supabase.auth.signUp(payload),
};

const supabaseMembers = {
  fetchProfile: (userId) => supabase.from('members').select('*').eq('user_id', userId).single(),
};

const supabaseAppData = {
  loadAll: () => supabase.from('app_data').select('key, value'),
  saveKey: (key, value) => supabase.from('app_data').upsert({ key, value, updated_at: new Date().toISOString() }),
};

const supabaseNotifications = {
  list: (limit = 100) => supabase.from('notifications').select('*').order('created_at', { ascending: false }).limit(limit),
  insert: (message, type = 'info') => supabase.from('notifications').insert({ message, type }),
  deleteAll: () => supabase.from('notifications').delete().gte('id', 0),
  deleteOne: (id) => supabase.from('notifications').delete().eq('id', id),
  subscribeInsert: (onInsert) => {
    const channel = supabase
      .channel('vmd-notifications')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'notifications' }, onInsert)
      .subscribe();
    return () => supabase.removeChannel(channel);
  },
};

const supabaseMail = {
  sendEmail: ({ to, subject, html }) => supabase.functions.invoke('send-email', { body: { to, subject, html } }),
};

const SIAN_BUCKET = 'sian-images';

const supabaseStorage = {
  // 시안 이미지를 base64로 JSON에 박아넣지 않고 Storage 버킷에 파일로 올린 뒤 공개 URL만 반환한다.
  uploadImage: async (file, folder = 'sian') => {
    const ext = (file.name.split('.').pop() || 'png').toLowerCase();
    const path = `${folder}/${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`;
    const { error } = await supabase.storage.from(SIAN_BUCKET).upload(path, file, {
      cacheControl: '31536000',
      contentType: file.type || undefined,
    });
    if (error) throw error;
    const { data } = supabase.storage.from(SIAN_BUCKET).getPublicUrl(path);
    return { url: data.publicUrl, path };
  },
  deleteImage: async (path) => {
    if (!path) return;
    await supabase.storage.from(SIAN_BUCKET).remove([path]);
  },
};

// ═══════════════════════════════════════════════════════════════════════
// Playground(vmddashboard) 경로 — vmddashboard-svc 백엔드(FastAPI + MariaDB) 사용.
// ═══════════════════════════════════════════════════════════════════════
const API_BASE = '/vmddashboard/api';

async function apiFetch(path, options = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    credentials: 'include',
    headers: { 'Content-Type': 'application/json' },
    ...options,
  });
  let body = null;
  try { body = await res.json(); } catch { /* no body */ }
  return { res, body };
}

// Supabase의 auth.onAuthStateChange 계약(콜백 등록 → 세션 변경 시 호출)을 흉내내는
// 아주 작은 로컬 이벤트 버스. 이 백엔드는 실시간 auth 이벤트 스트림이 없으므로,
// signInWithPassword/signOut이 직접 구독자에게 알려주는 방식으로 대체한다.
const authListeners = new Set();
function notifyAuthListeners(session) {
  authListeners.forEach((cb) => cb(session ? 'SIGNED_IN' : 'SIGNED_OUT', session));
}

// 가장 최근에 확인된 로그인 사용자 프로필 (fetchProfile이 재사용).
let cachedProfile = null;

function sessionFromUser(user) {
  if (!user) return null;
  cachedProfile = user;
  return { user: { id: user.id } };
}

const playgroundAuth = {
  getSession: async () => {
    const { res, body } = await apiFetch('/auth/session');
    if (res.status === 200 && body?.user) {
      return { data: { session: sessionFromUser(body.user) } };
    }
    cachedProfile = null;
    return { data: { session: null } };
  },
  onAuthStateChange: (cb) => {
    authListeners.add(cb);
    return { data: { subscription: { unsubscribe: () => authListeners.delete(cb) } } };
  },
  signOut: async () => {
    await apiFetch('/auth/logout', { method: 'POST' });
    cachedProfile = null;
    notifyAuthListeners(null);
    return { error: null };
  },
  signInWithPassword: async ({ email, password }) => {
    const { res, body } = await apiFetch('/auth/login', {
      method: 'POST',
      body: JSON.stringify({ email, password }),
    });
    if (res.status !== 200 || !body?.user) {
      return { error: { message: body?.detail || '이메일 또는 비밀번호가 올바르지 않습니다.' } };
    }
    notifyAuthListeners(sessionFromUser(body.user));
    return { error: null };
  },
  // Playground 배포본은 회원가입을 지원하지 않음 (관리자만 최초 업로드/계정 생성 가능) — 호출되지 않아야 함.
  signUp: async () => ({ error: { message: '이 배포본에서는 회원가입을 지원하지 않습니다.' } }),
};

const playgroundMembers = {
  fetchProfile: async (userId) => {
    if (cachedProfile && cachedProfile.id === userId) {
      return { data: { employee_id: cachedProfile.id, role: cachedProfile.role, name: cachedProfile.name } };
    }
    return { data: null };
  },
};

const playgroundAppData = {
  loadAll: async () => {
    const { res, body } = await apiFetch('/app-data');
    if (res.status !== 200) return { data: null };
    return { data: body.data };
  },
  saveKey: async (key, value) => {
    const { res, body } = await apiFetch(`/app-data/${encodeURIComponent(key)}`, {
      method: 'PUT',
      body: JSON.stringify({ value }),
    });
    if (res.status !== 200) return { error: body?.detail || 'save_failed' };
    return { error: null };
  },
};

const playgroundNotifications = {
  list: async (limit = 100) => {
    const { res, body } = await apiFetch(`/notifications?limit=${limit}`);
    if (res.status !== 200) return { data: null };
    return { data: body.data };
  },
  insert: async (message, type = 'info') => {
    const { res, body } = await apiFetch('/notifications', {
      method: 'POST',
      body: JSON.stringify({ message, type }),
    });
    if (res.status !== 200) return { error: body?.detail || 'insert_failed' };
    return { error: null };
  },
  deleteAll: async () => {
    await apiFetch('/notifications', { method: 'DELETE' });
  },
  deleteOne: async (id) => {
    await apiFetch(`/notifications/${id}`, { method: 'DELETE' });
  },
  // Realtime 대신 10초 폴링. 반환값은 Supabase의 removeChannel과 같은 역할(구독 해제 함수).
  subscribeInsert: (onInsert) => {
    let lastSeenIds = new Set();
    let initialized = false;
    const poll = async () => {
      const { data } = await playgroundNotifications.list(100);
      if (!data) return;
      if (!initialized) {
        // 첫 폴링은 "새 알림"으로 취급하지 않고 기존 목록만 기억한다.
        lastSeenIds = new Set(data.map((n) => n.id));
        initialized = true;
        return;
      }
      data.forEach((n) => {
        if (!lastSeenIds.has(n.id)) {
          lastSeenIds.add(n.id);
          onInsert({ new: n });
        }
      });
    };
    poll();
    const intervalId = setInterval(poll, 10000);
    return () => clearInterval(intervalId);
  },
};

// Playground Object Storage에 vmddashboard-svc(FastAPI)가 프록시로 업로드/삭제한다.
// uploadImage가 null을 반환하면 호출부(SianPopup)가 기존 base64 dataURL 저장 방식으로 폴백한다.
const playgroundStorage = {
  uploadImage: async (file) => {
    const formData = new FormData();
    formData.append('file', file);
    const res = await fetch(`${API_BASE}/files`, {
      method: 'POST',
      credentials: 'include',
      body: formData,
    });
    if (res.status !== 200) return null;
    const body = await res.json();
    return { url: body.url, path: body.path };
  },
  deleteImage: async (path) => {
    if (!path) return;
    await fetch(`${API_BASE}/files?path=${encodeURIComponent(path)}`, {
      method: 'DELETE',
      credentials: 'include',
    });
  },
};

const playgroundMail = {
  sendEmail: async ({ to, subject, html }) => {
    const { res, body } = await apiFetch('/send-email', {
      method: 'POST',
      body: JSON.stringify({ to, subject, html }),
    });
    if (res.status !== 200) return { error: body?.detail || 'send_failed' };
    return { error: null };
  },
};

// ═══════════════════════════════════════════════════════════════════════
export const dataClient = {
  auth: isVmddashboard ? playgroundAuth : supabaseAuth,
  members: isVmddashboard ? playgroundMembers : supabaseMembers,
  appData: isVmddashboard ? playgroundAppData : supabaseAppData,
  notifications: isVmddashboard ? playgroundNotifications : supabaseNotifications,
  mail: isVmddashboard ? playgroundMail : supabaseMail,
  storage: isVmddashboard ? playgroundStorage : supabaseStorage,
};

// 특정 키 하나의 "지금 서버에 있는" 값만 조회한다. 여러 세션이 같은 app_data 키를 동시에
// 쓸 때, 각자 로그인 시점 스냅샷을 그대로 덮어써서 서로의 변경을 지워버리는 문제를 막기 위해
// 저장 직전에 최신 값을 확인하고 병합하는 용도로 쓴다. 두 백엔드 모두 키별 GET이 없어서
// loadAll을 재사용해 필터링한다 (약간 비효율적이지만 백엔드 변경 없이 동작).
export async function loadAppDataKey(key) {
  const { data } = await dataClient.appData.loadAll();
  if (!data) return undefined;
  const row = data.find(r => r.key === key);
  return row ? row.value : undefined;
}
