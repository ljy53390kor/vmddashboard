export const hostname = typeof window !== 'undefined' ? window.location.hostname : '';
export const pathname = typeof window !== 'undefined' ? window.location.pathname : '';
export const isPG = hostname.includes('playground.idcube.sktelecom.com') || hostname.includes('playground.sktelecom.com');

// vmddashboard 앱 (Playground 배포본 — 자체 백엔드(vmddashboard-svc)를 통해 로그인/데이터 저장)
export const isVmddashboard = isPG && pathname.startsWith('/vmddashboard');
