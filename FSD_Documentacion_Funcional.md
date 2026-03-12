# Documento de Especificacion Funcional (FSD)

Proyecto analizado: `avaluos`  
Base del analisis: `README.md`, `requirements.txt`, `app.py`, `script.py`, `templates/index.html`, `static/app.js`, `docker-compose.yml`, `nginx/nginx.conf`.

## 1. Resumen Ejecutivo y Proposito
La aplicacion resuelve un problema operativo concreto: transformar informacion financiera extraida por IA (en formato JSON) en un archivo Excel de valuacion estandar, pasando antes por una revision humana guiada.

Valor entregado al usuario final:
- Reduce reprocesos al convertir automaticamente datos financieros a una plantilla corporativa.
- Permite corregir manualmente datos sensibles (por periodo y por rubro) antes de la salida final.
- Disminuye riesgo de errores de captura al centralizar validacion, previsualizacion y generacion en un mismo flujo.

Problema de negocio que resuelve:
- Los datos extraidos por IA pueden venir incompletos, con signos inconsistentes o con diferencias contables.
- El equipo de avaluos necesita estandarizar rapidamente esos datos en una hoja de trabajo unica para continuar el analisis financiero.

Proposito funcional del sistema:
- Ser una estacion de control "Humano + IA" para revisar, ajustar y publicar un Excel final de valuacion listo para uso interno.

## 2. Perfiles de Usuario (Roles y Permisos)
No existe autenticacion interna ni control de permisos por perfil dentro de la aplicacion. En la practica, el sistema opera con acceso uniforme para quien ingrese a la interfaz.

Perfiles funcionales observados:

1. Analista de Avaluos
- Puede cargar JSON (archivo o texto pegado).
- Puede editar metadata y cifras por anio.
- Puede agregar o eliminar periodos.
- Puede previsualizar la hoja consolidada.
- Puede generar y descargar el Excel final.

2. Responsable Operativo (mismo acceso funcional en pantalla)
- Tiene las mismas capacidades del analista.
- Adicionalmente, fuera de pantalla, puede definir que plantilla usar en despliegue.

3. Administracion Tecnica (infraestructura)
- Publica el servicio y configura proxy/reinicio.
- No hay permisos diferenciales en la interfaz; su rol es de operacion de entorno.

Implicacion funcional:
- Cualquier usuario con acceso al sitio tiene poder total sobre carga, correccion y generacion de resultados.

## 3. Inventario de Funcionalidades (Core Features)

### Modulo A. Carga de datos financieros
Descripcion funcional:
- El usuario incorpora informacion de entrada en JSON desde archivo local o pegado manual.

Flujo logico:
1. El usuario selecciona archivo o pega contenido.
2. El sistema valida estructura minima.
3. Si es valido, se crea un estado de trabajo editable con datos normalizados.
4. Si no es valido, se informa el error y no se habilita el flujo siguiente.

### Modulo B. Normalizacion de estructura y saneamiento inicial
Descripcion funcional:
- El sistema estandariza el contenido para evitar vacios, tipos inconsistentes y formatos mixtos.

Flujo logico:
1. Se completa una estructura base por periodo.
2. Valores vacios o no numericos se convierten a valores neutrales.
3. Se ordenan los periodos por anio.
4. Se sincroniza la lista de periodos detectados en metadata.

### Modulo C. Edicion guiada de metadata y cifras por periodo
Descripcion funcional:
- El usuario revisa campo por campo: identificadores temporales, estado de resultados, balance general y alerta analitica.

Flujo logico:
1. El sistema presenta tarjetas por anio con formularios guiados.
2. Cada cambio actualiza el estado de trabajo en pantalla.
3. Se recalcula el estado de equilibrio contable mostrado en el resumen de cada periodo.

### Modulo D. Gestion de periodos
Descripcion funcional:
- El usuario puede ampliar o depurar la serie historica para ajustar el alcance del analisis.

Flujo logico:
1. "Agregar anio" crea un nuevo periodo con estructura completa.
2. "Eliminar anio" quita el periodo seleccionado.
3. Si se elimina el ultimo periodo, el sistema crea automaticamente uno nuevo para mantener continuidad operativa.

### Modulo E. Resumen de alertas por anio
Descripcion funcional:
- Consolida observaciones cualitativas por periodo para revision final antes de exportar.

Flujo logico:
1. El usuario carga o edita texto de alerta por periodo.
2. El sistema compila esas alertas en un bloque resumen ordenado por anio.

### Modulo F. Previsualizacion operativa de hoja de trabajo
Descripcion funcional:
- El usuario revisa como quedaran los bloques principales del archivo final (Estado de Resultados y Balance General).

Flujo logico:
1. El sistema toma el estado actual validado.
2. Construye una vista tabular por anio.
3. Se muestran totales, separadores por seccion y coherencia de rubros para verificacion previa.

### Modulo G. Generacion y descarga de Excel final
Descripcion funcional:
- Ejecuta el llenado de plantilla y entrega un archivo final descargable.

Flujo logico:
1. El usuario confirma datos.
2. El sistema toma una plantilla Excel activa.
3. Inyecta la informacion por rubro y por periodo.
4. Genera archivo de salida con nombre trazable por fecha/hora.
5. Entrega descarga directa al usuario.

### Modulo H. Resolucion automatica de plantilla
Descripcion funcional:
- Define que plantilla usar sin pedir seleccion manual en cada ejecucion.

Flujo logico:
1. Si existe configuracion explicita, se usa esa plantilla.
2. Si no existe, se busca una plantilla valida en el directorio.
3. Si hay varias, se prioriza la plantilla institucional esperada.

## 4. Diccionario de Reglas de Negocio (Critico)
Reglas extraidas de la logica vigente:

1. El contenido de entrada debe ser un objeto JSON valido; de lo contrario, se rechaza.
2. La seccion `metadata` es obligatoria y debe tener formato de objeto.
3. La seccion `datos_financieros` es obligatoria y debe contener al menos un periodo.
4. El tamano maximo permitido para carga es 6 MB.
5. Si se define una plantilla explicita y no existe, la operacion se detiene con error.
6. Si no hay plantilla explicita, se busca automaticamente un archivo Excel en el proyecto.
7. Si hay varias plantillas candidatas, se prioriza la que coincide con la convencion "Grupo Ovando".
8. Si no existe ninguna plantilla Excel disponible, no se puede continuar con la generacion.
9. Un periodo marcado como parcial se etiqueta en salida como "`YYYY (Parcial)`"; en caso contrario se trata como anio cerrado.
10. Si un valor financiero viene vacio, nulo o invalido, se transforma a 0 para mantener consistencia operativa.
11. Los periodos se ordenan por anio para mantener secuencia cronologica estable.
12. Si sobran periodos respecto de capacidad de plantilla, se recorta la carga a la capacidad disponible.
13. Si faltan periodos respecto de capacidad de plantilla, las columnas libres se limpian con valores neutrales para no romper formulas.
14. El rubro "otros ingresos/gastos neto" se separa automaticamente: parte negativa como gasto y parte positiva como ingreso.
15. El total de "Pasivo + Capital" se calcula internamente como suma de ambos bloques, sin captura manual directa.
16. Al eliminar periodos, nunca se permite quedar sin ninguno: el sistema crea uno por defecto.
17. Al agregar un periodo, se propone como anio siguiente al mayor anio existente.
18. Si metadata no trae periodos detectados, el sistema los reconstruye a partir de los anios cargados.
19. En la vista de control, la ecuacion contable se considera balanceada cuando la diferencia es practicamente cero (tolerancia operativa minima).

## 5. Modelo de Datos Funcional
Entidades principales y relaciones logicas:

1. Metadata del Caso
- Identifica empresa, moneda y periodos detectados.
- Relacion: 1 Metadata agrupa N Periodos Financieros.

2. Periodo Financiero
- Unidad temporal de analisis (anual cerrado o parcial).
- Relacion: cada Periodo contiene 1 Estado de Resultados, 1 Balance General y 1 Alerta.

3. Estado de Resultados
- Representa desempeno acumulado del periodo: ventas, costos, utilidad bruta, gastos operativos, utilidad operativa, resultado financiero, impuestos, utilidad neta y depreciacion/amortizacion.
- Relacion: alimenta la seccion de resultados en la hoja final.

4. Balance General
- Representa posicion al cierre: activos (circulante y no circulante), pasivos (corto y largo plazo) y capital contable.
- Relacion: junto con Estado de Resultados soporta validaciones de coherencia contable.

5. Alerta Analitica por Periodo
- Texto de observaciones generado o ajustado por el usuario.
- Relacion: se consolida en un resumen transversal para control previo a exportacion.

6. Plantilla de Valuacion
- Estructura estandar de destino donde se cargan periodos y rubros.
- Relacion: recibe N Periodos y produce 1 archivo final.

7. Archivo Excel de Salida
- Entregable final para consumo del equipo de valuacion.
- Relacion: resultado de combinar Metadata + Periodos + reglas de mapeo hacia plantilla.

## 6. Integraciones Externas
Integraciones funcionales identificadas:

1. Servicio web de tipografia (Google Fonts)
- Uso funcional: estandarizar la presentacion visual de la interfaz de revision.

2. Recurso externo de icono corporativo
- Uso funcional: branding visual del producto en navegador.

3. Capa de publicacion con proxy web (Nginx)
- Uso funcional: exponer el sistema hacia usuarios finales y enrutar trafico al servicio de aplicacion.

4. Motor de ejecucion de servicio (Gunicorn en contenedor)
- Uso funcional: operar la aplicacion en entorno productivo de forma estable.

No se detectaron integraciones funcionales con pasarelas de pago, mensajeria, correo transaccional o servicios fiscales en el alcance actual.
