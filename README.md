# Informe de Remuneraciones

Pagina estatica publicada en GitHub Pages.

## Actualizar datos

Los libros fuente estan en `data`. Reemplaza el archivo del grupo correspondiente manteniendo su nombre y confirma el cambio en `main`:

- CRUX: `detalle_remuneraciones.xlsx`
- Grupo NOW: `now_detalle_remuneraciones.xlsx`
- Grupo AVESA: `avesa_detalle_remuneraciones.xlsx`
- Grupo Avanza: `avanza_libro_remuneraciones.xlsx`

GitHub Actions regenerara las paginas del informe automaticamente. Los meses disponibles dependen de los procesos incluidos en cada libro.

El libro completo es la fuente definitiva para CRUX. No se aplican sobrescrituras de archivos mensuales ni correcciones anteriores.

Cuando termine la accion, el mismo link de GitHub Pages mostrara los datos nuevos para todos.

El archivo de carga interno debe tener las cabeceras en la fila 5 y la hoja `Detalle`.

Para Grupo Avanza también se reconoce directamente el libro de remuneraciones con la hoja `Libro de remu CE` y sus cabeceras en la fila 6. La aplicación detecta ese formato automáticamente al seleccionarlo desde `Cargar archivo`.
