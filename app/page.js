"use client";
import { useState } from 'react';
import mammoth from 'mammoth';

export default function Home() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleProcess = async () => {
    if (!file) return alert("Por favor, sube un archivo Word primero");
    setLoading(true);
    
    try {
      console.log("⏳ Extrayendo texto con Mammoth en el navegador...");
      
      // 1. Convertimos y extraemos el texto del Word
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
      const textoExtraido = result.value;

      if (!textoExtraido) throw new Error("El documento parece estar vacío o no se pudo leer.");

      console.log("✅ Texto extraído. Enviando a la Inteligencia Artificial (Gemini)...");
      
      // 2. Llamamos a nuestro puente seguro en el backend
      const iaResponse = await fetch('/api/analizar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ texto: textoExtraido })
      });

      const iaData = await iaResponse.json();

      if (iaData.success) {
        console.log("🧠 Respuesta Estructurada de la IA:\n", iaData.datos);
        alert("¡La IA analizó el documento exitosamente! Revisa la consola del navegador para ver la magia estructurada en JSON.");
        
        // ¡El paso final (Fase 3) será usar la librería 'docx' para convertir este JSON en el Word final!
      } else {
        alert("Error de la IA: " + iaData.error);
      }

    } catch (error) {
      console.error("❌ Error:", error);
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
          {loading ? 'Analizando con IA...' : 'Aplicar Normas APA'}
        </button>
      </div>
    </main>
  );
}