# Catalyst · Quiminova — guía para publicarlo en la nube

Sigue estos pasos en orden. Son dos partes: primero la base de datos (Supabase), después la publicación web (Vercel). En total toma unos 15-20 minutos y ambos servicios son gratis para este tamaño de uso.

## Parte 1 · Base de datos (Supabase)

1. **Crea una cuenta gratis** en https://supabase.com (puedes entrar con tu cuenta de Google).
2. Clic en **"New project"**. Ponle un nombre (ej. `quiminova`), elige una contraseña para la base de datos (guárdala, no la necesitarás de nuevo) y elige la región más cercana (ej. `South America`). Espera 1-2 minutos a que se cree.
3. En el menú de la izquierda, entra a **"SQL Editor"** → **"New query"**.
4. Abre el archivo `supabase-schema.sql` que viene en esta carpeta, copia **todo** su contenido, pégalo ahí, y dale clic a **"Run"**. Esto crea la tabla donde vivirán tus datos, protegida para que solo tú puedas verlos.
5. Ve a **"Authentication" → "Users"** → clic en **"Add user"** → **"Create new user"**. Escribe tu correo y una contraseña (esta va a ser tu usuario y clave para entrar al sistema — no hay registro público, solo tú puedes entrar). Marca la casilla de **"Auto Confirm User"** si aparece.
6. Ve a **"Project Settings" (ícono de engranaje) → "API"**. Copia dos valores, los vas a necesitar en la Parte 2:
   - **Project URL**
   - **anon public** key (la clave larga que empieza con "eyJ...")

## Parte 2 · Publicar la app (Vercel)

1. **Sube esta carpeta a GitHub**: crea una cuenta gratis en https://github.com si no tienes, crea un repositorio nuevo (puede ser privado), y sube todos los archivos de esta carpeta (`quiminova-catalyst`). Si nunca lo has hecho, GitHub tiene un botón de "uploading an existing file" donde puedes simplemente arrastrar la carpeta.
2. Crea una cuenta gratis en https://vercel.com, entra con tu cuenta de GitHub.
3. Clic en **"Add New" → "Project"**, elige el repositorio que acabas de subir, y dale **"Import"**.
4. Antes de darle "Deploy", abre la sección **"Environment Variables"** y agrega estas dos (con los valores que copiaste en el paso 6 de la Parte 1):
   - `VITE_SUPABASE_URL` = tu Project URL
   - `VITE_SUPABASE_ANON_KEY` = tu anon public key
5. Clic en **"Deploy"**. En 1-2 minutos Vercel te da una dirección web, algo como `https://quiminova-catalyst.vercel.app` — esa es tu sistema, ya en la nube, disponible desde cualquier computador o celular.
6. Entra a esa dirección con el correo y contraseña que creaste en el paso 5 de la Parte 1.

Listo — desde ahí todo lo que hagas (ventas, pagos, inventario, etc.) se guarda automáticamente en tu base de datos de Supabase y no se pierde aunque cierres el navegador, cambies de computador o se dañe algo en Vercel.

## ¿Y si quiero mi propio dominio (ej. sistema.quiminova.com)?

En Vercel, entra al proyecto → **"Settings" → "Domains"** → agrega tu dominio y sigue las instrucciones (te pedirá agregar un registro DNS donde compraste el dominio).

## ¿Y si quiero probarlo primero en mi computador?

Necesitas tener [Node.js](https://nodejs.org) instalado. Luego, en esta carpeta:

```
npm install
cp .env.example .env
```

Edita `.env` con tus valores de Supabase, y luego:

```
npm run dev
```

Se abrirá en `http://localhost:5173`.

## Notas importantes

- **Solo tú puedes entrar**: no hay pantalla de registro. Si en el futuro alguien más de tu equipo necesita entrar, créale un usuario nuevo desde Supabase (Authentication → Users → Add user) — aunque ese usuario vería una base de datos vacía separada de la tuya, porque cada usuario tiene su propio set de datos. Dime si en algún momento quieres que varias personas compartan la misma información y te ayudo a ajustarlo.
- **Respaldo**: Supabase hace respaldos automáticos de tu base de datos en su plan gratuito, pero solo por un tiempo limitado. Si tu negocio crece, vale la pena revisar los planes pagos de Supabase (son baratos) para respaldos más largos.
- **Costo**: mientras el uso sea el de un solo negocio, tanto Supabase como Vercel se mantienen en su capa gratuita.
