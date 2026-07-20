import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://ukhhrxrwaoqwmamkcxgu.supabase.co';
const supabaseKey =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVraGhyeHJ3YW9xd21hbWtjeGd1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM1MjA5OTMsImV4cCI6MjA5OTA5Njk5M30.LDaaadDwTYxRjP4VI_Y0OIuBZrNCpi_3TQIcBhibpO8';

export const supabase = createClient(supabaseUrl, supabaseKey);
