# PLANI PRESENTACION LEVEL

Aplicación Streamlit para generar presentaciones educativas. Las cuentas se
gestionan con Supabase Auth. No hay registro público: el administrador crea
las cuentas de los clientes y ellos inician sesión con sus credenciales.

## Configuración de Supabase

1. Crea un proyecto en [Supabase](https://supabase.com/) o usa uno existente.
2. En el panel del proyecto, habilita el proveedor de autenticación **Email**.
3. Copia la **Project URL** y la clave pública **anon** (o **publishable**).
   No uses la `service_role` ni una clave secreta en la aplicación.
4. En Streamlit Community Cloud, abre la aplicación, entra a **Settings >
   Secrets** y agrega:

   ```toml
   SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
   SUPABASE_ANON_KEY = "TU-CLAVE-PUBLICA"
   ```

5. Guarda los cambios y reinicia la aplicación. Para desarrollo local, coloca
   esas mismas variables en `.streamlit/secrets.toml`; ese archivo está
   excluido de Git.

## Crear cuentas de clientes sin enviar correos

En **Authentication > Users**, crea cada usuario directamente desde el panel
de Supabase y establece su contraseña. Confirma la dirección del correo al
crear la cuenta (`email_confirm = true` si utilizas la Admin API). La creación
directa con `create_user` no envía correo de confirmación ni invitación.
Comparte las credenciales con el cliente de forma privada.

La aplicación solo permite iniciar sesión. Nunca configures la clave
`service_role` en Streamlit Secrets ni en el código de la aplicación.
