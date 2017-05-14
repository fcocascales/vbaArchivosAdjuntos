# Adjuntar ficheros a una tabla desde un formulario de MS-Access

Refactorización del módulo de [neckkito](http://siliconproject.com.ar/neckkito/)

## Código

Es un módulo de clase **ClaseArchivosAdjuntos.cls** para una base de datos Microsoft Access.

## Características

- Flexibilidad en el nombre de los campos
- Facilitar al máximo su uso desde el formulario
- Poder ser usado en formularios no independientes
- Renombrar variables
- Creación de nuevas funciones para no repetir código
- Reescribir comentarios y eliminar otros
- Eliminar variables globales usadas como locales
- Facilitar renombrado de las carpetas
- Reescritura de código para facilitar su compresión
- Convertirlo en un módulo de Clase
- Añadir propiedad *CarpetaDatos* para separar datos por tablas
- Añadir propiedad *NombreTabla* para indicar la ruta raíz a la BD de tabla vinculada
- Añadir propiedad *CarpetaRaiz* para ubicar exactamente los adjuntos

## Requisitos

Registrar estas librerías (>Herramientas>Referencias)

 - Visual Basic For Applications
 - Microsoft Access xx.x Object Library
 - Microsoft Office xx.x Access database engine Object Library // Microsoft DAO Object Library
 - Microsoft Office xx.x Object Library
 - Microsoft Scripting Runtime

## Forma de empleo

### Tabla

En la tabla se necesita que exista:

 - un campo único **id**
 - y un campo de tipo *memo* o *texto largo* **adjuntos**

### Formulario

#### Se necesitan 3 controles en el formulario:

  1. Un id que debe ser único y es como se llamarán las subcarpetas dentro e Adjuntos/Datos
  2. Un campo para guardar el nombre de los adjuntos. Debe ser de tipo texto. Debería ser visible=no
  3. Un combobox independiente para seleccionar el documento adjunto actual

#### Variables del formulario:

	  private archivosAdjuntos as ClasearchivosAdjuntos

#### Eventos de formulario:

 - Form_Load=>

		Set archivosAdjuntos = New ClaseArchivosAdjuntos
		archivosAdjuntos.NombreTabla = "Datos"
		archivosAdjuntos.CampoId = me.id
		archivosAdjuntos.CampoAdjuntos = me.adjuntos
		archivosAdjuntos.ComboboxIndependiente = me.comboDocumento
		archivosAdjuntos.Inicializar

 - Form_AfterInsert=>

		Call archivosAdjuntos.Procesar

 - Form_AfterUpdate=>

		Call archivosAdjuntos.Procesar

 - Form_Current=>

		Call archivosAdjuntos.Actualizar
		Call archivosAdjuntos.Asear

#### Botones del formulario:

 - Boton1=>

       Call archivosAdjuntos.BotonElegir

 - Boton2=>

       Call archivosAdjuntos.BotonVer

 - Boton3=>

       Call archivosAdjuntos.BotonExportar

 - Boton4=>

       Call archivosAdjuntos.BotonBorrar
