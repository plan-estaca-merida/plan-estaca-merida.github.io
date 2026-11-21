# Énfasis 2026 · Estaca de Sión Mérida

Página de seguimiento del Énfasis del Área del Caribe 2026 ("Ministrar uno a uno a través de ordenanzas y convenios") para los barrios de la Estaca de Sión Mérida.

- Sitio: https://plan-estaca-merida.github.io/
- Por ahora solo muestra el Barrio El Castor (modo prueba). Los demás barrios están ocultos, no borrados.
- Datos compartidos: hoja de Google "Énfasis 2026 · Barrio El Castor (datos)" a través de `apps-script/Codigo.gs` (aplicación web). La URL `/exec` va en `var API` dentro de `index.html`. Mientras `API` esté vacío, cada quien guarda solo en su propio teléfono.
- Código de acceso: solo el consejo de barrio entra. El script exige el código en cada lectura y cada guardado; abrir la URL `/exec` sin código no muestra nada. El código no está en el repositorio: vive en las Propiedades del script (`CODIGO`) y se crea o cambia ejecutando `nuevoCodigo` en el editor de Apps Script. Se comparte por WhatsApp en privado con el enlace `https://plan-estaca-merida.github.io/#acceso=XXXX-XXXX-XXXX`; el teléfono lo guarda y lo quita de la barra de direcciones. Si el código cambia, los teléfonos con el viejo borran los datos guardados y vuelven a pedirlo.
- Uso simultáneo: cada cambio guarda solo los campos tocados y la hoja los combina uno a la vez, así dos líderes pueden trabajar al mismo tiempo sin borrarse lo que anotó el otro. La página revisa cambios cada 8 segundos.

Herramienta local de apoyo. No es un documento oficial de La Iglesia de Jesucristo de los Santos de los Últimos Días.
