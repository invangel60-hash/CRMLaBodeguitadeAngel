# La Bodeguita de Angel — Web Shopping Cart

## 🚀 Inicio rápido

1. **Configura Supabase**:
   - Abre `index.html` y busca el bloque `// ── Configuration ──`
   - Reemplaza `SUPABASE_URL` y `SUPABASE_ANON_KEY` con los valores de tu proyecto

   ```js
   const SUPABASE_URL      = 'https://TU-PROYECTO.supabase.co';
   const SUPABASE_ANON_KEY = 'TU_ANON_KEY';
   ```

2. **Abre el archivo en el navegador** — no necesita servidor, funciona con `file://` o cualquier hosting estático.

---

## 🗄️ Estructura de tablas Supabase esperada

```sql
-- Clasificaciones
clasificacion_articulos (id_clasificacion, nombre)

-- Artículos
articulo (id_art, nombre, stock, precio, id_clasificacion, permite_decimales, inf)

-- Pedidos (cabecera)
pedidos (id_pedido, fecha, nro_tlf, nombre_cliente, status)

-- Detalle del carrito
pedidos_renglones (id_renglon, id_pedido, id_articulo, cantidad, precio)
```

---

## ✨ Características

| Funcionalidad                        | Estado |
|--------------------------------------|--------|
| Listado de productos desde Supabase  | ✅     |
| Filtro por clasificación (tabs)      | ✅     |
| Solo muestra productos con stock > 0 | ✅     |
| Control de cantidad (entero/decimal) | ✅     |
| Carrito lateral deslizable           | ✅     |
| Ajuste de cantidad dentro del carrito| ✅     |
| Modal de pedido (nombre + teléfono)  | ✅     |
| Guardado en `pedidos`                | ✅     |
| Guardado en `pedidos_renglones`      | ✅     |
| Modal de confirmación con # de pedido| ✅     |
| Toast de notificación                | ✅     |
| Diseño responsivo mobile-first       | ✅     |
| Skeleton loaders                     | ✅     |
| Sanitización XSS                     | ✅     |

---

## ⚙️ Supabase Row Level Security (RLS)

Asegúrate de que las políticas RLS de Supabase permitan:
- `SELECT` en `articulo` y `clasificacion_articulos` (anon)
- `INSERT` en `pedidos` y `pedidos_renglones` (anon)

```sql
-- Ejemplo de política mínima
CREATE POLICY "Allow public read on articulo"
  ON articulo FOR SELECT TO anon USING (true);

CREATE POLICY "Allow public insert on pedidos"
  ON pedidos FOR INSERT TO anon WITH CHECK (true);

CREATE POLICY "Allow public insert on pedidos_renglones"
  ON pedidos_renglones FOR INSERT TO anon WITH CHECK (true);
```

---

## 📁 Estructura del proyecto

```
CRM/
├── index.html   ← Aplicación completa (HTML + CSS + JS)
├── .env         ← Placeholder para tus credenciales (no incluir en git)
└── README.md
```
