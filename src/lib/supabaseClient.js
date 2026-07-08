import { createClient } from '@supabase/supabase-js';

// Vercel(및 그 외 기본) 배포본 전용 — Supabase Cloud 그대로 사용.
// Playground(vmddashboard) 배포본은 dataClient.js를 통해 별도 백엔드(vmddashboard-svc)를 사용하므로
// 이 클라이언트를 참조하지 않는다.
const EXTERNAL_URL = 'https://ytevgwijapjhwifqiaco.supabase.co';
const EXTERNAL_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl0ZXZnd2lqYXBqaHdpZnFpYWNvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk3NzUwMzgsImV4cCI6MjA5NTM1MTAzOH0.6YDJz3XoUqYee4nGRPG0OdA3n8zEdQSeiZwajY9dq1o';

export const supabase = createClient(EXTERNAL_URL, EXTERNAL_KEY);
