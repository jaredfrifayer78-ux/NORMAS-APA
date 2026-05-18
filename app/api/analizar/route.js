import { GoogleGenerativeAI } from '@google/generative-ai';
import { NextResponse } from 'next/server';

export async function POST(request) {
  try {
    const { texto } = await request.json();

    if (!texto) {
      return NextResponse.json({ error: "No hay texto para analizar" }, { status: 400 });
    }

    // Inicializamos Gemini con tu clave secreta
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    // Usamos el modelo Flash porque es rapidísimo y maneja documentos largos
    const model = genAI.getGenerativeModel({ 
      model: "gemini-1.5-flash-latest",
      generationConfig: { responseMimeType: "application/json" } // Obligamos a que devuelva JSON
    });

    const promptMaestro = `
    Eres un sistema experto en análisis estructural de textos académicos bajo Normas APA 7ma edición.
    Tu única tarea es analizar el texto, clasificar cada fragmento y devolver un JSON estructurado.
    
    REGLAS:
    1. NO alteres, resumas ni corrijas el texto original.
    2. Divide el texto proporcionado en bloques lógicos.
    3. Asigna a cada bloque uno de estos tipos: "portada_titulo", "portada_autor", "portada_institucion", "titulo_nivel_1", "titulo_nivel_2", "parrafo", "cita_bloque", "referencia".
    
    TEXTO A ANALIZAR:
    ${texto}
    
    FORMATO DE SALIDA ESPERADO (JSON Puro):
    {
      "documento": [
        { "tipo": "parrafo", "contenido": "Texto aquí..." }
      ]
    }
    `;

    console.log("Enviando texto a Gemini...");
    const result = await model.generateContent(promptMaestro);
    const response = await result.response;
    const jsonIA = JSON.parse(response.text());

    console.log("¡Análisis de IA completado!");
    return NextResponse.json({ success: true, datos: jsonIA });

  } catch (error) {
    console.error("Error en la IA:", error);
    return NextResponse.json({ error: "Fallo al comunicar con la IA" }, { status: 500 });
  }
}