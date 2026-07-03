import { createClient } from '@supabase/supabase-js';

const hostname = typeof window !== 'undefined' ? window.location.hostname : '';
const pathname = typeof window !== 'undefined' ? window.location.pathname : '';
const isPG = hostname.includes('playground.idcube.sktelecom.com') || hostname.includes('playground.sktelecom.com');

// vmddata 앱 (내부망 Supabase)
const isVmddata     = isPG && pathname.startsWith('/vmddata');
// vmddashboard 앱 (외부 Supabase 프록시 시도)
const isVmddashboard = isPG && pathname.startsWith('/vmddashboard');

const EXTERNAL_URL = 'https://ytevgwijapjhwifqiaco.supabase.co';
const EXTERNAL_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl0ZXZnd2lqYXBqaHdpZnFpYWNvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk3NzUwMzgsImV4cCI6MjA5NTM1MTAzOH0.6YDJz3XoUqYee4nGRPG0OdA3n8zEdQSeiZwajY9dq1o';

const INTERNAL_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiYW5vbiIsImlzcyI6InN1cGFiYXNlIiwiaWF0IjoxNTE2MjM5MDIyLCJleHAiOjE4OTM0NTYwMDB9.jQSaJ1e7ETayvxn2I0P1nFVCsTl2Sa-Nk4JaEGCcOtA';

let supabaseUrl, supabaseAnonKey;
if (isVmddashboard) {
  supabaseUrl     = `${window.location.origin}/vmddashboard/supabase-api`;
  supabaseAnonKey = INTERNAL_KEY || EXTERNAL_KEY;
} else {
  supabaseUrl     = EXTERNAL_URL;
  supabaseAnonKey = EXTERNAL_KEY;
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
