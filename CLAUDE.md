# CLAUDE.md — Registro de Auditorías y Base de OPs (lives-registro-auditorias)

Las reglas generales de trabajo están en el repo privado `lives-instrucciones` (CLAUDE.md). Este archivo solo describe esta herramienta.

## Qué hace
Calidad Manufactura registra aquí sus auditorías. Además, este repo mantiene la base de OPs y los programas de PCP, que usan otras herramientas.
- `registro.html` — **formulario** "Registro de auditorías".
  - Tipos LINEAS, SERVINSP, EMPAQUE, CLIENTE y FINAL. `?tipo=` fija el tipo.
  - Carga individual o por hoja de varias filas.
  - También busca, corrige y anula registros.
- `catalogos.html` — **mantenimiento** de las listas: personal, talleres, destinos, clientes, líneas y defectos.
- `ops.html` — **carga** del Excel de OPs exportado del sistema de producción (columnas `Cod_GrupoPro`, `Est_Prop`, `OP`, `fec_creacion_op`…). Desde el 27/10/2026 las OPs que vienen en el archivo se reemplazan y las que no vienen se MANTIENEN: una carga nunca borra OPs de la tabla `ops` (ya no hay aviso de "faltan").
- `consulta-ops.html` — **consulta** "Base de OPs", con Excel.
- `partidas.html` — **carga** de la base de partidas (desde el 23/10/2026): primera pestaña del Excel "Partidas" del sistema ("Desagrupado por OP y OC"). Columnas con los nombres de la fila 1 (PARTIDA, CODIGO DE TELA, NOMBRE DE TELA, CODIGO DE COLOR, NOMBRE DE COLOR, GRUPO, PESO PROGRAMADO, OP, OC); acepta también los del sistema de la fila 2 (`cod_ordprov`, `cod_tela`, `des_tela`, `cod_color`, `des_color`, `cod_grupotex`, `Peso_Progr`, `OP`, `Orden_Corte`) y salta esa segunda fila si viene. Desde el 27/10/2026 las partidas que vienen en el archivo se reemplazan y las que no vienen se mantienen (sin aviso de "faltan"). Sirve para cruzar con otras bases por OP (el grupo es el textil y puede no ser el de la OP, ej. `RES-THY`).
- `prendas-por-pedido.html` — **análisis** de prendas promedio por pedido según cliente, tipo de pedido, grupo, estilo, etc. Tiene Excel.
- `programas/subir.html` — **carga** de los programas de PCP:
  - Costura: archivo .xlsm, pestaña DESTINO.
  - Acabados: archivo .xlsx, pestaña SERVICIOS.
  - El programa anterior queda guardado.
- `programas/index.html` — **consulta** "Programas de Costura y Acabados". Tiene Excel.

## Base de datos (Supabase lives-datos)
- Auditorías (`registro.html`):
  - Guardar: `registro_guardar_directo`, `registro_corregir_directo` y `registro_anular_directo`. Escriben en `auditorias`, `auditoria_defectos`, `auditoria_colores`, `auditoria_auditoras` y `auditoria_destinos`. Son las mismas tablas de `lives-auditoria-final`.
  - Consultar: `registro_buscar`, `registro_buscaraud`, `registro_leer`, `registro_recientes` y `registro_sugerir`.
  - `es_tipo_directo` lee `auditorias_ajustes`.
  - `sync_agregar_cat` agrega un valor al catálogo sobre la marcha.
  - Catálogos: vistas `publico_ref_*`, que salen de las tablas `cat_*` y de `ops`.
- Catálogos (`catalogos.html`): lee directo `cat_personal`, `cat_talleres`, `cat_destinos`, `cat_clientes`, `cat_lineas` y `cat_defectos`. Cambia con `sync_cat_nuevo`, `sync_cat_editar`, `sync_cat_activar` y `sync_defecto_editar`.
- OPs:
  - `ops.html` usa `sync_ops_reemplazar`, que reemplaza solo las OPs que vienen en el archivo y mantiene las demás (BD03).
  - Las consultas leen `publico_ref_ops` y la fecha de la última copia en `publico_estado_copia` (de `sync_estado`).
  - La tabla `ops` la usan también Auditoría Final, Captura de Auditorías de Taller y otras.
- Partidas: `partidas.html` usa `sync_partidas_reemplazar` (security definer, carga libre para pruebas firmando con el nombre; supervisión no guarda), que reemplaza solo las partidas que vienen en el archivo, mantiene las demás (BD03) y anota la carga en `sync_estado` (fila `partidas`). Lectura libre. Registro: `BD02 base de partidas`.
- Programas:
  - `prog_validar_ops` y `prog_publicar` escriben en `programas_filas`, `programas_filas_ant`, `programas_cols`, `programas_publicaciones` y `sync_estado`.
  - Las consultas leen las vistas `publico_programas` y `publico_programas_estado`.
- Este repo no usa Storage.

## Quién la usa
Según las entradas por el portal de los últimos 30 días al 21/10/2026: 13 personas. Son auditoras de calidad (7), la digitadora de Calidad Manufactura, supervisoras de calidad y de Top Sample, y gerencia.

## Quién puede guardar
- Catálogos y carga de OPs: solo el área de Datos e IA (`es_superusuario()` en la base).
- Carga de partidas: libre mientras se prueba (cualquiera, escribiendo su nombre). En SQL Server, Sistemas exige usuario.
- Publicar programas: lo decide la función de permiso `prog_puede_subir`, con la tabla `permisos`.
- Registrar auditorías: la OP tiene que existir en la base y la auditora tiene que estar en el catálogo de personal.
- A Sistemas: `prog_puede_subir` y `es_superusuario` son las piezas que reemplazan con su seguridad, junto con su control de roles (ver o registrar).

## Pendientes vistos al escribir esto (21/10/2026)
- `registro.html` lee la vista `publico_alias_personal`, que no existe en la base.
- Quedan nombres heredados del Apps Script: la pasarela `jsonp(...)` con acciones guardar, editar y anular. Si se llega por ese camino, sale "Esta acción ya no está disponible".
