"use client";
import { useState } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';

export default function Home() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  // Función para construir el Word final usando las reglas APA
  const generarDocumentoAPA = async (datosIA) => {
    console.log("⏳ Ensamblando el nuevo documento Word...");
    
    // Recorremos el JSON y creamos un párrafo de Word por cada bloque
    const parrafosWord = datosIA.documento.map(bloque => {
      
      // Configuración base para todo el documento: Times New Roman 12, interlineado doble
      let config = {
        children: [new TextRun({ text: bloque.contenido, font: "Times New Roman", size: 24 })], // size 24 = 12pt en Word
        spacing: { line: 480 }, // 480 es interlineado doble exacto
      };

      // Aplicamos las reglas APA según el tipo de bloque que detectó la IA
      switch(bloque.tipo) {
        case 'portada_titulo':
        case 'titulo_nivel_1':
          config.alignment = AlignmentType.CENTER;
          config.children[0] = new TextRun({ text: bloque.contenido, font: "Times New Roman", size: 24, bold: true });
          break;
        case 'titulo_nivel_2':
          config.alignment = AlignmentType.LEFT;
          config.children[0] = new TextRun({ text: bloque.contenido, font: "Times New Roman", size: 24, bold: true });
          break;
        case 'parrafo':
          config.alignment = AlignmentType.LEFT;
          config.indent = { firstLine: 720 }; // Sangría de primera línea de 0.5 pulgadas
          break;
        case 'cita_bloque':
          config.indent = { left: 720 }; // Sangría izquierda completa de 0.5 pulgadas
          break;
        case 'referencia':
          config.indent = { left: 720, hanging: 720 }; // Sangría francesa clásica de bibliografías
          break;
        default: // portada_autor, portada_institucion...
          config.alignment = AlignmentType.CENTER;
          break;
      }
      return new Paragraph(config);
    });

    // Creamos el documento aplicando márgenes APA (1 pulgada = 1440 twips en todos los lados)
    const doc = new Document({
      sections: [{
        properties: {
          page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
        },
        children: parrafosWord
      }]
    });

    // Empaquetamos y forzamos la descarga del archivo en el navegador
    const blob = await Packer.toBlob(doc);
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Documento_APA_Formateado.docx";
    a.click();
    window.URL.revokeObjectURL(url);
    
    console.log("✅ ¡Documento generado y descargado!");
  };

  const handleProcess = async () => {
    if (!file) return alert("Por favor, sube un archivo Word primero");
    setLoading(true);
    
    try {
      // FASE 1: Extracción
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
      const textoExtraido = result.value;
      if (!textoExtraido) throw new Error("Documento vacío");

      // FASE 2: Análisis Estructural con IA
      const iaResponse = await fetch('/api/analizar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ texto: textoExtraido })
      });

      const iaData = await iaResponse.json();

      if (iaData.success) {
        // FASE 3: Generación del nuevo Word
        await generarDocumentoAPA(iaData.datos);
        alert("¡Misión cumplida! Revisa tu carpeta de descargas.");
      } else {
        alert("Error de la IA: " + iaData.error);
      }

    } catch (error) {
      console.error("❌ Error general:", error);
      alert("Hubo un problema: " + error.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen bg-gray-50 flex flex-col items-center justify-center p-6">
      <div className="max-w-md w-full bg-white rounded-xl shadow-md p-8 text-center">
        <h1 className="text-2xl font-bold text-gray-800 mb-2">Formateador APA 7ma</h1>
        <p className="text-gray-500 mb-6">Sube tu documento y deja que la IA organice la estructura por ti.</p>
        
        <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 mb-6 bg-gray-50">
          <input 
            type="file" 
            accept=".docx" 
            onChange={handleFileChange}
            className="block w-full text-sm text-gray-500
              file:mr-4 file:py-2 file:px-4
              file:rounded-md file:border-0
              file:text-sm file:font-semibold
              file:bg-blue-100 file:text-blue-700
              hover:file:bg-blue-200 cursor-pointer"
          />
        </div>

        <button 
          onClick={handleProcess}
          disabled={!file || loading}
          className={`w-full py-3 px-4 rounded-lg text-white font-medium transition-colors ${
            !file || loading 
              ? 'bg-gray-400 cursor-not-allowed' 
              : 'bg-blue-600 hover:bg-blue-700 shadow-md'
          }`}
        >
          {loading ? 'Aplicando Normas APA...' : 'Aplicar Normas APA'}
        </button>
      </div>
    </main>
  );
}