import { createClient } from '@supabase/supabase-js';

const isPlayground = typeof window !== 'undefined' &&
  (window.location.hostname.includes('playground.idcube.sktelecom.com') ||
   window.location.hostname.includes('playground.sktelecom.com'));

const supabaseUrl = isPlayground
  ? `${window.location.origin}/vmddashboard/supabase-api`
  : 'https://ytevgwijapjhwifqiaco.supabase.co';

const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl0ZXZnd2lqYXBqaHdpZnFpYWNvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk3NzUwMzgsImV4cCI6MjA5NTM1MTAzOH0.6YDJz3XoUqYee4nGRPG0OdA3n8zEdQSeiZwajY9dq1o';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
