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

      DriveApp.getFolderById(CONFIG.CARPETA_RAIZ_ID);



    const properties =

      PropertiesService.getScriptProperties();



    const registros =

      properties.getProperties();



    const idsActuales = new Set();



    // Se calcula una sola vez por ejecución.

    let maxCodigo =

      obtenerMaximoCodigo_(datos, cCodigo);





    // ─────────────────────────────────────

    // 4. PROCESAR OBRAS

    // ─────────────────────────────────────



    for (let i = 1; i < datos.length; i++) {



      const fila = datos[i];



      // Ignorar filas incompletas.

      if (!fila[cID] || !fila[cTitulo]) {

        continue;

      }



      const obraID = String(fila[cID]);

      idsActuales.add(obraID);





      // ─────────────────────────────────

      // CÓDIGO

      // ─────────────────────────────────



      let codigo = fila[cCodigo];



      if (!codigo) {



        maxCodigo++;



        codigo =

          'OBR-' +

          String(maxCodigo).padStart(6, '0');



        hoja

          .getRange(i + 1, cCodigo + 1)

          .setValue(codigo);



      }





      // ─────────────────────────────────

      // PERSONAS

      // ─────────────────────────────────



      const compositor =

        resolverValorRef_(

          fila[cCompositor],

          mapaPersonas

        );



      const arreglador =

        resolverValorRef_(

          fila[cArreglador],

          mapaPersonas

        );





      // ─────────────────────────────────

      // NOMBRE CORRECTO DE CARPETA

      // ─────────────────────────────────



      let nombreCarpeta = fila[cTitulo];



      if (compositor) {

        nombreCarpeta += ' (' + compositor + ')';

      }



      nombreCarpeta += ' - ' + codigo;



      nombreCarpeta =

        limpiarNombreArchivo_(nombreCarpeta);





      // ─────────────────────────────────

      // CREAR / RECUPERAR CARPETA

      // ─────────────────────────────────



      let carpetaObra;



      if (fila[cCarpetaID]) {



        try {



          carpetaObra =

            DriveApp.getFolderById(

              String(fila[cCarpetaID])

            );



          // Renombrar únicamente si cambió

          // Título o Compositor.

          if (

            carpetaObra.getName() !==

            nombreCarpeta

          ) {



            carpetaObra.setName(nombreCarpeta);



          }





          // AUTORREPARACIÓN:

          // si existe la carpeta pero falta su URL,

          // reconstruirla automáticamente.

          if (!fila[cCarpetaURL]) {



            hoja

              .getRange(i + 1, cCarpetaURL + 1)

              .setValue(carpetaObra.getUrl());



          }



        } catch (error) {



          console.error(

            'No se pudo acceder a la carpeta de ' +

            codigo +

            ': ' +

            error

          );



          continue;



        }



      } else {



        // No existe carpeta todavía:

        // crearla.

        carpetaObra =

          carpetaRaiz.createFolder(

            nombreCarpeta

          );



        const nuevoCarpetaID =

          carpetaObra.getId();



        const nuevaCarpetaURL =

          carpetaObra.getUrl();



        // Guardar ID.

        hoja

          .getRange(i + 1, cCarpetaID + 1)

          .setValue(nuevoCarpetaID);



        // Guardar URL.

        hoja

          .getRange(i + 1, cCarpetaURL + 1)

          .setValue(nuevaCarpetaURL);



      }





      // ─────────────────────────────────

      // REGISTRAR OBRA → CARPETA

      // ─────────────────────────────────



      const clave =

        CONFIG.PREFIJO_REGISTRO + obraID;



      const carpetaID =

        carpetaObra.getId();



      // Solo escribir Properties

      // cuando sea necesario.

      if (registros[clave] !== carpetaID) {



        properties.setProperty(

          clave,

          carpetaID

        );



        registros[clave] = carpetaID;



      }





      // ─────────────────────────────────

      // ARCHIVO NUEVO DE APPSHEET

      // ─────────────────────────────────



      const referenciaArchivo =

        fila[cArchivo];



      if (referenciaArchivo) {



        moverArchivoAppSheet_(

          ss,

          referenciaArchivo,

          carpetaObra,

          codigo,

          fila[cTitulo],

          compositor

        );



      }





      // ─────────────────────────────────

      // FICHA TÉCNICA PDF

      // ─────────────────────────────────



      try {



        const datosFicha = {



          TITULO:

            String(fila[cTitulo] || ''),



          COMPOSITOR:

            compositor,



          ARREGLADOR:

            arreglador,



          GENERO:

            resolverListaRef_(

              fila[cGenero],

              mapaGeneros

            ),



          FORMACION_VOCAL:

            resolverListaRef_(

              fila[cFormacion],

              mapaFormaciones

            ),



          INSTRUMENTACION:

            String(

              fila[cInstrumentacion] || ''

            ),



          TIEMPO_LITURGICO:

            resolverListaRef_(

              fila[cTiempo],

              mapaTiempos

            ),



          USO_LITURGICO:

            resolverListaRef_(

              fila[cUso],

              mapaUsos

            ),



          CODIGO:

            String(codigo || ''),



          MATERIAL:

            carpetaObra.getUrl(),



          NOTAS:

            String(fila[cNotas] || '')



        };





        generarFichaTecnicaSiHaceFalta_(

          obraID,

          carpetaObra,

          fila[cTitulo],

          compositor,

          datosFicha,

          properties

        );



      } catch (error) {



        // Un problema en la ficha técnica

        // NO debe detener el resto del sistema.

        console.error(

          'No se pudo generar la ficha técnica de ' +

          codigo +

          ': ' +

          error

        );



      }



    }





    // ─────────────────────────────────────

    // 5. DETECTAR OBRAS BORRADAS

    // ─────────────────────────────────────



    detectarObrasBorradas_(

      idsActuales,

      registros,

      properties

    );





  } catch (error) {



    console.error(

      'Error general: ' + error

    );



  } finally {



    lock.releaseLock();



  }



}





// ───────────────────────────────────────────

// GENERAR / ACTUALIZAR FICHA TÉCNICA

// ───────────────────────────────────────────



function generarFichaTecnicaSiHaceFalta_(

  obraID,

  carpetaDestino,

  titulo,

  compositor,

  datosFicha,

  properties

) {



  const claveHash =

    CONFIG.PREFIJO_FICHA_HASH + obraID;



  const contenidoHash =

    JSON.stringify(datosFicha);



  const hashNuevo =

    Utilities.base64EncodeWebSafe(

      Utilities.computeDigest(

        Utilities.DigestAlgorithm.SHA_256,

        contenidoHash,

        Utilities.Charset.UTF_8

      )

    );





  // Nombre definitivo del PDF.

  let nombrePDF = titulo;



  if (compositor) {

    nombrePDF += ' (' + compositor + ')';

  }



  nombrePDF += ' - Ficha técnica.pdf';



  nombrePDF =

    limpiarNombreArchivo_(nombrePDF);





  // Buscar si ya existe una ficha técnica.

  const fichaExistente =

    buscarFichaTecnica_(carpetaDestino);





  const hashAnterior =

    properties.getProperty(claveHash);





  // Si los datos no cambiaron y la ficha existe

  // con el nombre correcto, no hacer nada.

  if (

    hashAnterior === hashNuevo &&

    fichaExistente &&

    fichaExistente.getName() === nombrePDF

  ) {

    return;

  }





  // ─────────────────────────────────────

  // CREAR COPIA TEMPORAL DE LA PLANTILLA

  // ─────────────────────────────────────



  const plantilla =

    DriveApp.getFileById(

      CONFIG.PLANTILLA_FICHA_ID

    );



  const copiaTemporal =

    plantilla.makeCopy(

      'TEMP Ficha técnica - ' + obraID

    );



  const idTemporal =

    copiaTemporal.getId();





  try {



    const documento =

      DocumentApp.openById(idTemporal);



    const cuerpo =

      documento.getBody();





    // Reemplazar marcadores.

    reemplazarMarcador_(

      cuerpo,

      '{{TITULO}}',

      datosFicha.TITULO

    );



    reemplazarMarcador_(

      cuerpo,

      '{{COMPOSITOR}}',

      datosFicha.COMPOSITOR

    );



    reemplazarMarcador_(

      cuerpo,

      '{{ARREGLADOR}}',

      datosFicha.ARREGLADOR

    );



    reemplazarMarcador_(

      cuerpo,

      '{{GENERO}}',

      datosFicha.GENERO

    );



    reemplazarMarcador_(

      cuerpo,

      '{{FORMACION_VOCAL}}',

      datosFicha.FORMACION_VOCAL

    );



    reemplazarMarcador_(

      cuerpo,

      '{{INSTRUMENTACION}}',

      datosFicha.INSTRUMENTACION

    );



    reemplazarMarcador_(

      cuerpo,

      '{{TIEMPO_LITURGICO}}',

      datosFicha.TIEMPO_LITURGICO

    );



    reemplazarMarcador_(

      cuerpo,

      '{{USO_LITURGICO}}',

      datosFicha.USO_LITURGICO

    );



    reemplazarMarcador_(

      cuerpo,

      '{{CODIGO}}',

      datosFicha.CODIGO

    );



    reemplazarMarcador_(

      cuerpo,

      '{{MATERIAL}}',

      datosFicha.MATERIAL

    );



    reemplazarMarcador_(

      cuerpo,

      '{{NOTAS}}',

      datosFicha.NOTAS

    );





    documento.saveAndClose();





    // Convertir a PDF.

    const pdfBlob =

      DriveApp

        .getFileById(idTemporal)

        .getAs(MimeType.PDF)

        .setName(nombrePDF);





    // Si ya existe la ficha, actualizar su contenido
    // conservando el mismo archivo y el mismo ID de Drive.
    // Si todavía no existe, crearla normalmente.

    if (fichaExistente) {

      Drive.Files.update(
        { name: nombrePDF },
        fichaExistente.getId(),
        pdfBlob,
        { fields: 'id,name' }
      );

    } else {

      carpetaDestino.createFile(pdfBlob);

    }





    // Guardar hash únicamente si todo salió bien.

    properties.setProperty(

      claveHash,

      hashNuevo

    );





  } finally {



    // La copia de Google Docs es solo temporal.

    // Se envía siempre a Papelera.

    try {



      DriveApp

        .getFileById(idTemporal)

        .setTrashed(true);



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
