const CONFIG = {



  HOJA_OBRAS: 'OBRAS',

  HOJA_PERSONAS: 'PERSONAS',



  // Carpeta 01. PARTITURAS

  CARPETA_RAIZ_ID: '14Bw1XLgi3f3K6LMtvFCa8Eek50fneWn7',



  // Plantilla — Ficha técnica Biblioteca LAUS

  PLANTILLA_FICHA_ID: '1pI6CC52EWl7m3UzdioAHRaelCE18_yXo16zW-vEt82k',



  PREFIJO_REGISTRO: 'OBRA_',

  PREFIJO_FICHA_HASH: 'FICHA_HASH_'



};





function procesarObrasNuevas() {



  const lock = LockService.getScriptLock();



  // Evita ejecuciones simultáneas.

  if (!lock.tryLock(10000)) {

    console.log('Ya hay otra ejecución en curso.');

    return;

  }



  try {



    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hoja = ss.getSheetByName(CONFIG.HOJA_OBRAS);



    if (!hoja) {

      throw new Error('No existe la hoja OBRAS.');

    }





    // ─────────────────────────────────────

    // 1. LEER OBRAS UNA SOLA VEZ

    // ─────────────────────────────────────



    const datos = hoja.getDataRange().getValues();



    if (datos.length < 1) {

      return;

    }



    const encabezados = datos[0];

    const col = nombre => encabezados.indexOf(nombre);



    const cID = col('ID');

    const cCodigo = col('Código');

    const cTitulo = col('Título');

    const cCompositor = col('Compositor');

    const cArreglador = col('Arreglador');

    const cTiempo = col('Tiempo litúrgico');

    const cUso = col('Uso litúrgico');

    const cFormacion = col('Formación vocal');

    const cInstrumentacion = col('Instrumentación');

    const cGenero = col('Género');

    const cArchivo = col('Archivo');

    const cCarpetaID = col('Carpeta Drive ID');

    const cCarpetaURL = col('Carpeta Drive URL');

    const cNotas = col('Notas');





    // Verificación de seguridad.

    const columnasNecesarias = {



      ID: cID,

      Código: cCodigo,

      Título: cTitulo,

      Compositor: cCompositor,

      Arreglador: cArreglador,

      'Tiempo litúrgico': cTiempo,

      'Uso litúrgico': cUso,

      'Formación vocal': cFormacion,

      Instrumentación: cInstrumentacion,

      Género: cGenero,

      Archivo: cArchivo,

      'Carpeta Drive ID': cCarpetaID,

      'Carpeta Drive URL': cCarpetaURL,

      Notas: cNotas



    };



    for (const nombre in columnasNecesarias) {



      if (columnasNecesarias[nombre] === -1) {



        throw new Error(

          'Falta la columna "' + nombre + '" en OBRAS.'

        );



      }



    }





    // ─────────────────────────────────────

    // 2. MAPAS DE AUTORIDADES

    // ─────────────────────────────────────



    const mapaPersonas =

      crearMapaTabla_(ss, CONFIG.HOJA_PERSONAS);



    const mapaFormaciones =

      crearMapaTabla_(ss, 'Formación vocal');



    const mapaTiempos =

      crearMapaTabla_(ss, 'Tiempo litúrgico');



    const mapaUsos =

      crearMapaTabla_(ss, 'Uso litúrgico');



    const mapaGeneros =

      crearMapaTabla_(ss, 'Género');





    // ─────────────────────────────────────

    // 3. PREPARACIÓN

    // ─────────────────────────────────────



    const carpetaRaiz =

      // El documento temporal no forma parte del archivo:
    // eliminarlo definitivamente para no generar residuos.
    Drive.Files.remove(idTemporal);



    } catch (error) {



      console.error(

        'No se pudo eliminar la copia temporal de la ficha: ' +

        error

      );



    }



  }



}





// ───────────────────────────────────────────

// BUSCAR FICHA TÉCNICA EXISTENTE

// ───────────────────────────────────────────



function buscarFichaTecnica_(carpeta) {



  const archivos =

    carpeta.getFiles();



  while (archivos.hasNext()) {



    const archivo =

      archivos.next();



    const nombre =

      archivo.getName();



    if (

      / - Ficha técnica\\.pdf$/i.test(nombre)

    ) {



      return archivo;



    }



  }



  return null;



}





// ───────────────────────────────────────────

// REEMPLAZAR MARCADOR EN GOOGLE DOC

// ───────────────────────────────────────────



function reemplazarMarcador_(

  cuerpo,

  marcador,

  valor

) {



  // replaceText usa expresiones regulares,

  // por eso escapamos el marcador.

  const patron =

    marcador.replace(

      /[.\*+?^${}()|[\\]\\\\]/g,

      '\\\\$&'

    );



  cuerpo.replaceText(

    patron,

    String(valor || '')

  );



}





// ───────────────────────────────────────────

// RESOLVER UN VALOR REF

// ───────────────────────────────────────────



function resolverValorRef_(

  valor,

  mapa

) {



  if (

    valor === null ||

    valor === undefined ||

    valor === ''

  ) {

    return '';

  }



  const texto =

    String(valor).trim();



  return mapa.get(texto) || texto;



}





// ───────────────────────────────────────────

// RESOLVER ENUM / ENUMLIST DE REF

// ───────────────────────────────────────────



function resolverListaRef_(

  valor,

  mapa

) {



  if (

    valor === null ||

    valor === undefined ||

    valor === ''

  ) {

    return '';

  }



  let valores;



  if (Array.isArray(valor)) {



    valores = valor;



  } else {



    let texto =

      String(valor).trim();



    // AppSheet suele guardar EnumList

    // como una lista separada por comas.

    valores =

      texto

        .split(',')

        .map(v => v.trim())

        .filter(Boolean);



  }



  return valores

    .map(v => mapa.get(String(v)) || String(v))

    .join(', ');



}





// ───────────────────────────────────────────

// DETECTAR OBRAS BORRADAS

// ───────────────────────────────────────────



function detectarObrasBorradas_(

  idsActuales,

  registros,

  properties

) {



  for (const clave in registros) {



    if (

      !clave.startsWith(

        CONFIG.PREFIJO_REGISTRO

      )

    ) {

      continue;

    }



    const obraID =

      clave.substring(

        CONFIG.PREFIJO_REGISTRO.length

      );



    // La obra sigue existiendo.

    if (idsActuales.has(obraID)) {

      continue;

    }



    const carpetaID =

      registros[clave];



    try {



      const carpeta =

        DriveApp.getFolderById(carpetaID);



      // Seguridad:

      // no eliminar definitivamente.

      // Solo enviar a Papelera.

      carpeta.setTrashed(true);



      // El registro interno se elimina

      // únicamente después de que Drive

      // completó la operación.

      properties.deleteProperty(clave);



      // Limpiar también el hash

      // de la ficha técnica.

      properties.deleteProperty(

        CONFIG.PREFIJO_FICHA_HASH + obraID

      );



      console.log(

        'Obra eliminada. Carpeta enviada a Papelera: ' +

        carpetaID

      );



    } catch (error) {



      // Si Drive falla, conservar el registro

      // para volver a intentarlo.

      console.error(

        'No se pudo enviar la carpeta a Papelera: ' +

        carpetaID +

        ' — ' +

        error

      );



    }



  }



}





// ───────────────────────────────────────────

// CREAR MAPA DE TABLA

// ───────────────────────────────────────────



function crearMapaTabla_(

  ss,

  nombreHoja

) {



  const hoja =

    ss.getSheetByName(nombreHoja);



  if (!hoja) {



    throw new Error(

      'No existe la hoja "' +

      nombreHoja +

      '".'

    );



  }



  const datos =

    hoja.getDataRange().getValues();



  const mapa = new Map();



  if (datos.length < 2) {

    return mapa;

  }



  const encabezados = datos[0];



  const cID =

    encabezados.indexOf('ID');



  const cNombre =

    encabezados.indexOf('Nombre');



  if (cID === -1 || cNombre === -1) {



    throw new Error(

      '"' +

      nombreHoja +

      '" debe contener las columnas ID y Nombre.'

    );



  }



  for (let i = 1; i < datos.length; i++) {



    if (datos[i][cID]) {



      mapa.set(

        String(datos[i][cID]),

        datos[i][cNombre]

      );



    }



  }



  return mapa;



}





// ───────────────────────────────────────────

// CREAR MAPA DE PERSONAS

// Conservada por compatibilidad.

// ───────────────────────────────────────────



function crearMapaPersonas_(ss) {



  return crearMapaTabla_(

    ss,

    CONFIG.HOJA_PERSONAS

  );



}





// ───────────────────────────────────────────

// OBTENER MAYOR CÓDIGO EXISTENTE

// ───────────────────────────────────────────



function obtenerMaximoCodigo_(

  datos,

  indiceCodigo

) {



  let maximo = 0;



  for (let i = 1; i < datos.length; i++) {



    const coincidencia =

      String(

        datos[i][indiceCodigo] || ''

      ).match(/^OBR-(\d+)$/);



    if (coincidencia) {



      maximo =

        Math.max(

          maximo,

          Number(coincidencia[1])

        );



    }



  }



  return maximo;



}





// ───────────────────────────────────────────

// MOVER ARCHIVO NUEVO DE APPSHEET

// ───────────────────────────────────────────



function moverArchivoAppSheet_(

  ss,

  referenciaArchivo,

  carpetaDestino,

  codigo,

  titulo,

  compositor

) {



  const partes =

    String(referenciaArchivo).split('/');



  if (partes.length < 2) {

    return;

  }



  const nombreCarpetaOrigen =

    partes[0];



  const nombreArchivoOriginal =

    partes[partes.length - 1];





  // Conservar extensión original.

  const punto =

    nombreArchivoOriginal.lastIndexOf('.');



  const extension =

    punto >= 0

      ? nombreArchivoOriginal.substring(punto)

      : '';





  // El archivo recibe este nombre

  // SOLAMENTE cuando entra por primera vez.

  let nombreNuevo = titulo;



  if (compositor) {

    nombreNuevo += ' - ' + compositor;

  }



  nombreNuevo +=

    ' - ' + codigo + extension;



  nombreNuevo =

    limpiarNombreArchivo_(nombreNuevo);





  // Encontrar la carpeta Base de datos.

  const archivoSpreadsheet =

    DriveApp.getFileById(ss.getId());



  const padres =

    archivoSpreadsheet.getParents();



  if (!padres.hasNext()) {

    return;

  }



  const carpetaBaseDatos =

    padres.next();





  // Encontrar carpeta temporal creada

  // por AppSheet.

  const carpetasOrigen =

    carpetaBaseDatos.getFoldersByName(

      nombreCarpetaOrigen

    );



  if (!carpetasOrigen.hasNext()) {

    return;

  }



  const carpetaOrigen =

    carpetasOrigen.next();





  // Encontrar el archivo original.

  const archivos =

    carpetaOrigen.getFilesByName(

      nombreArchivoOriginal

    );



  if (!archivos.hasNext()) {



    // Normalmente significa que el archivo

    // ya fue procesado y movido.

    return;



  }



  const archivo =

    archivos.next();





  // Renombrar solo en esta primera entrada.

  archivo.setName(nombreNuevo);



  // Mover a la carpeta definitiva.

  archivo.moveTo(carpetaDestino);



}





// ───────────────────────────────────────────

// LIMPIAR NOMBRES

// ───────────────────────────────────────────



function limpiarNombreArchivo_(nombre) {



  return String(nombre)

    .replace(/[\\/\\\\:\*?"<>|]/g, '-')

    .replace(/\s+/g, ' ')

    .trim();



}
