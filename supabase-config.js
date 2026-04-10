// Supabase Configuration for PC Manager
const SUPABASE_URL = 'https://yyknaylktzanlcnwuecb.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl5a25heWxrdHphbmxjbnd1ZWNiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNDcwMjQsImV4cCI6MjA4OTkyMzAyNH0.9U0mO5suI6z7LVeSYliRV78QjSiV6xcEvaJc8SprCk0';

// Use var to make client globally accessible (overwrites library namespace intentionally)
var supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
