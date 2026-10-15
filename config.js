// Configuración de Supabase (Project Settings → API).
// La anon key es pública por diseño; el acceso lo controlan las políticas RLS de supabase/setup.sql.
// Si se dejan vacíos, la app funciona en "modo local" (sin guardar nada).
window.APP_CONFIG = {
  SUPABASE_URL: "https://mlmsympewmxfcqoyowxs.supabase.co",
  SUPABASE_ANON_KEY: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1sbXN5bXBld214ZmNxb3lvd3hzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk3Njk0MDEsImV4cCI6MjEwNTM0NTQwMX0.G4D-QkugQ5rMP_rXj4WGf2SaE8mv305QHjPe-171Srg",
  BUCKET: "evidencias"
};
