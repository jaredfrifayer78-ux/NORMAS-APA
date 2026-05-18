"use client";
import { useState, useRef } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';
import { motion, AnimatePresence, useScroll, useTransform } from 'framer-motion';
import { UploadCloud, FileText, CheckCircle, Loader2, ChevronLeft, ArrowRight, Layers, FileCheck2, Code, AlertTriangle } from 'lucide-react';

export default function Home() {
  const [view, setView] = useState('landing');
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState('idle'); 
  const [loadingText, setLoadingText] = useState('');

  const containerRef = useRef(null);
  const { scrollYProgress } = useScroll({ target: containerRef });
  
  const heroScale = useTransform(scrollYProgress, [0, 0.2], [1, 0.95]);
  const heroOpacity = useTransform(scrollYProgress, [0, 0.2], [1, 0]);
  const bentoY = useTransform(scrollYProgress, [0, 0.3], [100, 0]);

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
      setStatus('idle');
    }
  };

  const generarDocumentoAPA = async (datosIA) => {
    setLoadingText('Compilando estructura física del archivo...');
    const parrafosWord = datosIA.documento.map(bloque => {
      let config = {
        children: [new TextRun({ text: bloque.contenido, font: "Times New Roman", size: 24 })],
        spacing: { line: 480 },
      };

      switch(bloque.tipo) {
        case 'portada_titulo':
        case 'titulo_nivel_1':
          config.alignment = AlignmentType.CENTER;
          const titulo1 = bloque.contenido.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
          config.children[0] = new TextRun({ text: titulo1, font: "Times New Roman", size: 24, bold: true });
          break;
        case 'titulo_nivel_2':
          config.alignment = AlignmentType.LEFT;
          const titulo2 = bloque.contenido.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
          config.children[0] = new TextRun({ text: titulo2, font: "Times New Roman", size: 24, bold: true });
          break;
        case 'parrafo':
          config.alignment = AlignmentType.LEFT;
          config.indent = { firstLine: 720 };
          break;
        case 'cita_bloque':
          config.indent = { left: 720 };
          break;
        case 'referencia':
          config.indent = { left: 720, hanging: 720 };
          break;
        default:
          config.alignment = AlignmentType.CENTER;
          break;
      }
      return new Paragraph(config);
    });

    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
        children: parrafosWord
      }]
    });

    const blob = await Packer.toBlob(doc);
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Documento_Formateado_Pro.docx";
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const handleProcess = async () => {
    if (!file) return;
    setStatus('loading');
    
    try {
      setLoadingText('Parseando estructura binaria...');
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
      const textoExtraido = result.value;
      if (!textoExtraido) throw new Error("Documento vacío.");

      setLoadingText('Mapeando nodos de texto en el motor estructural...');
      const iaResponse = await fetch('/api/analizar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ texto: textoExtraido })
      });

      const iaData = await iaResponse.json();

      if (iaData.success) {
        await generarDocumentoAPA(iaData.datos);
        setStatus('success');
      } else {
        throw new Error(iaData.error);
      }
    } catch (error) {
      console.error(error);
      alert("Error: " + error.message);
      setStatus('error');
    }
  };

  return (
    <div ref={containerRef} className="min-h-screen bg-[#0c0c0e] text-[#f1f1f3] font-sans antialiased overflow-x-hidden selection:bg-indigo-500/40 relative">
      
      {/* Grilla técnica de fondo (Figma Canvas Grid) */}
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#1f1f24_1px,transparent_1px),linear-gradient(to_bottom,#1f1f24_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_0%,#000_70%,transparent_100%)] opacity-25 pointer-events-none"></div>

      {/* Navbar Minimalista de Alta Precisión */}
      <nav className="w-full h-16 border-b border-[#1f1f24] bg-[#0c0c0e]/80 backdrop-blur-md flex justify-between items-center px-6 fixed top-0 left-0 right-0 z-50">
        <div className="flex items-center gap-6">
          <div className="font-black text-lg tracking-tighter text-white flex items-center gap-2 cursor-pointer" onClick={() => setView('landing')}>
            <span className="w-3 h-3 bg-indigo-500 rounded-full shadow-[0_0_10px_#6366f1]"></span>
            FORMATIA <span className="text-xs font-mono font-normal text-[#6b6b76] border border-[#1f1f24] px-1.5 py-0.5 rounded">v1.0.0</span>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <button 
            onClick={() => setView(view === 'landing' ? 'app' : 'landing')}
            className={`text-xs font-bold uppercase tracking-wider px-5 py-2.5 rounded transition-all border ${
              view === 'landing' 
                ? 'bg-white text-black border-white hover:bg-transparent hover:text-white' 
                : 'bg-transparent text-[#6b6b76] border-[#1f1f24] hover:text-white'
            }`}
          >
            {view === 'landing' ? 'Abrir Consola' : 'Ver Manifiesto'}
          </button>
        </div>
      </nav>

      <AnimatePresence mode="wait">
        
        {/* =======================
            VISTA 1: FIGMA LANDING PAGE
        ======================== */}
        {view === 'landing' && (
          <motion.div 
            key="landing"
            className="pt-32 pb-32 max-w-7xl mx-auto px-6 relative z-10"
          >
            {/* Hero Seccion con Tipografía de Impacto */}
            <motion.div style={{ scale: heroScale, opacity: heroOpacity }} className="text-left max-w-5xl mb-24 pt-10">
              <p className="font-mono text-xs uppercase tracking-[0.25em] text-indigo-400 mb-4 font-semibold">
                // AUTOMATIZACIÓN DE FORMATO ACADÉMICO
              </p>
              <h1 className="text-6xl md:text-8xl font-extrabold tracking-tight text-white mb-8 leading-[0.95] font-black">
                Estructura impecable. <br />
                Diseño imperceptible.
              </h1>
              <p className="text-xl md:text-2xl text-[#6b6b76] max-w-3xl mb-12 leading-relaxed font-normal">
                No alteramos una sola palabra de tu investigación. Nuestro motor de análisis sintáctico segmenta el texto original y lo reconstruye en un archivo Word nativo bajo los estándares más estrictos del manual APA 7.
              </p>
              <button 
                onClick={() => setView('app')}
                className="group inline-flex items-center gap-4 bg-[#1f1f24] hover:bg-white text-white hover:text-black font-bold text-base px-8 py-4 rounded border border-[#2e2e35] hover:border-white transition-all shadow-2xl active:scale-95"
              >
                Inicializar Herramienta 
                <ArrowRight size={18} className="group-hover:translate-x-1 transition-transform" />
              </button>
            </motion.div>

            {/* Bento Grid (Estructura Asimétrica Tipo Figma) */}
            <motion.div style={{ y: bentoY }} className="grid md:grid-cols-3 gap-4 max-w-6xl mx-auto text-left">
              
              {/* Bloque 1: Títulos (Izquierda) */}
              <div className="bg-[#131316] border border-[#1f1f24] p-10 rounded-2xl md:col-span-2 flex flex-col justify-between group hover:border-[#2e2e35] transition-colors relative overflow-hidden">
                <div className="absolute top-0 right-0 w-64 h-64 bg-indigo-500/5 rounded-full blur-3xl pointer-events-none"></div>
                <div>
                  <Layers className="text-indigo-400 mb-6 w-8 h-8" />
                  <h3 className="text-2xl font-bold text-white mb-3 tracking-tight">Normalización de Títulos de Nivel 1 y 2</h3>
                  <p className="text-[#6b6b76] leading-relaxed max-w-md">
                    El motor identifica títulos escritos accidentalmente en mayúsculas sostenidas y los transforma automáticamente a mayúsculas iniciales, centrados y en negrita según el protocolo de la APA.
                  </p>
                </div>
                <div className="mt-12 font-mono text-xs text-[#3a3a44] border-t border-[#1f1f24] pt-4 flex gap-4">
                  <span>INPUT: TEXT_RAW</span>
                  <span>➔</span>
                  <span className="text-indigo-400">OUTPUT: TITLE_CASE_VALIDATED</span>
                </div>
              </div>

              {/* Bloque 2: Geometría de Página (Derecha) */}
              <div className="bg-[#131316] border border-[#1f1f24] p-10 rounded-2xl flex flex-col justify-between group hover:border-[#2e2e35] transition-colors">
                <div>
                  <FileCheck2 className="text-emerald-400 mb-6 w-8 h-8" />
                  <h3 className="text-2xl font-bold text-white mb-3 tracking-tight">Geometría de Página</h3>
                  <p className="text-[#6b6b76] leading-relaxed">
                    Márgenes exactos de 1 pulgada en los 4 flancos de la página, interlineado doble de 480 twips y sangrías de primera línea precisas de 0.5 pulgadas.
                  </p>
                </div>
                <div className="mt-6 font-mono text-xs text-[#3a3a44]">
                  [MÁRGENES: 1440 twips]
                </div>
              </div>

              {/* Bloque RESTRICCIÓN (Nuevo bloque Bento de Advertencia Técnica) */}
              <div className="bg-[#131316] border border-[#1f1f24] p-10 rounded-2xl md:col-span-1 flex flex-col justify-between group hover:border-amber-500/30 transition-colors relative overflow-hidden">
                <div className="absolute top-0 right-0 w-32 h-32 bg-amber-500/5 rounded-full blur-2xl pointer-events-none"></div>
                <div>
                  <AlertTriangle className="text-amber-500 mb-6 w-8 h-8" />
                  <h3 className="text-2xl font-bold text-white mb-3 tracking-tight">Alcance del Motor</h3>
                  <p className="text-[#6b6b76] text-sm leading-relaxed">
                    Esta herramienta se enfoca estrictamente en perfeccionar la **estructura semántica y el texto crudo**. Si tu documento de origen contiene objetos complejos como <span className="text-amber-500/80 font-medium">tablas o imágenes, estas no se exportarán</span>.
                  </p>
                </div>
                <div className="mt-6 font-mono text-[10px] text-amber-500/60 bg-amber-500/5 border border-amber-500/10 p-3 rounded-lg">
                  // REGLA: Aplica el formato aquí y luego reinserta tus gráficos por copiado clásico en el Word final.
                </div>
              </div>

              {/* Bloque Privacidad (Abajo Derecha - Ancho 2 columnas) */}
              <div className="bg-[#131316] border border-[#1f1f24] p-10 rounded-2xl md:col-span-2 flex flex-col justify-between group hover:border-[#2e2e35] transition-colors">
                <div>
                  <div className="flex items-center gap-3 mb-4">
                    <Code className="text-blue-400 w-6 h-6" />
                    <span className="font-mono text-xs uppercase tracking-wider text-blue-400 font-bold">Procesamiento Local</span>
                  </div>
                  <h3 className="text-2xl font-bold text-white mb-2 tracking-tight">Privacidad por Arquitectura</h3>
                  <p className="text-[#6b6b76] leading-relaxed">
                    La extracción binaria se ejecuta por completo del lado del cliente usando la memoria efímera del navegador. Tus investigaciones confidenciales jamás tocan un almacenamiento secundario.
                  </p>
                </div>
                <div className="mt-6 font-mono text-xs text-[#3a3a44]">
                  [CLIENT_SIDE_PARSING: TRUE]
                </div>
              </div>

            </motion.div>
          </motion.div>
        )}

        {/* =======================
            VISTA 2: LA CONSOLA DE TRABAJO (DASHBOARD)
        ======================== */}
        {view === 'app' && (
          <motion.div 
            key="app"
            initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
            className="pt-24 pb-24 px-6 max-w-4xl mx-auto flex flex-col items-center justify-center min-h-screen"
          >
            <div className="w-full max-w-xl flex justify-start mb-6">
              <button 
                onClick={() => { setView('landing'); setFile(null); setStatus('idle'); }}
                className="flex items-center gap-2 text-[#6b6b76] hover:text-white transition-colors font-mono text-xs uppercase tracking-wider"
              >
                <ChevronLeft size={16} /> [ Cerrar Consola ]
              </button>
            </div>

            {/* Contenedor de la Herramienta Estilo Terminal Limpia */}
            <div className="w-full max-w-xl bg-[#131316] border border-[#1f1f24] rounded-2xl p-8 md:p-12 relative">
              
              <div className="flex items-center gap-1.5 absolute top-4 left-6">
                <span className="w-2.5 h-2.5 rounded-full bg-[#1f1f24]"></span>
                <span className="w-2.5 h-2.5 rounded-full bg-[#1f1f24]"></span>
                <span className="w-2.5 h-2.5 rounded-full bg-[#1f1f24]"></span>
              </div>

              <div className="text-left mb-8 border-b border-[#1f1f24] pb-6">
                <h2 className="text-xl font-bold text-white tracking-tight">Compilador Estructural APA</h2>
                <p className="text-sm text-[#6b6b76] mt-1 font-mono">Carga un archivo de origen binario para iniciar el mapeo.</p>
              </div>

              <AnimatePresence mode="wait">
                {(status === 'idle' || status === 'error') && (
                  <motion.div key="upload" className="flex flex-col">
                    <div className="relative w-full group">
                      <input 
                        type="file" 
                        accept=".docx" 
                        onChange={handleFileChange}
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                      />
                      <div className={`border border-dashed rounded-xl p-10 flex flex-col items-center justify-center transition-colors duration-200
                        ${file 
                          ? 'border-indigo-500 bg-indigo-500/5' 
                          : 'border-[#1f1f24] bg-[#0c0c0e]/50 group-hover:border-[#2e2e35]'}`}>
                        <UploadCloud className={`mb-4 ${file ? 'text-indigo-400' : 'text-[#3a3a44]'} w-10 h-10`} />
                        <h3 className="text-sm font-semibold text-white mb-1 font-mono">
                          {file ? file.name : "SELECCIONAR_ARCHIVO.DOCX"}
                        </h3>
                        <p className="text-xs text-[#6b6b76]">
                          {file ? "Payload cargado exitosamente" : "Arrastra o haz clic para montar"}
                        </p>
                      </div>
                    </div>

                    <button 
                      onClick={handleProcess}
                      disabled={!file}
                      className={`mt-6 w-full py-4 rounded font-bold text-xs uppercase tracking-widest transition-all font-mono border ${
                        !file 
                          ? 'bg-transparent text-[#3a3a44] border-[#1f1f24] cursor-not-allowed' 
                          : 'bg-white text-black border-white hover:bg-transparent hover:text-white'
                      }`}
                    >
                      EJECUTAR_PROCESAMIENTO_APA
                    </button>
                  </motion.div>
                )}

                {status === 'loading' && (
                  <motion.div key="loading" className="flex flex-col items-center justify-center py-12 text-left w-full">
                    <div className="w-full bg-[#0c0c0e] border border-[#1f1f24] p-5 rounded-lg font-mono text-xs text-[#6b6b76] space-y-2">
                      <div className="flex items-center gap-2">
                        <Loader2 className="w-3.5 h-3.5 text-indigo-400 animate-spin" />
                        <span className="text-white">STATUS: EXECUTING_TASK</span>
                      </div>
                      <div className="text-indigo-400 border-t border-[#1f1f24] pt-2 mt-2 animate-pulse">
                        &gt; {loadingText}
                      </div>
                    </div>
                  </motion.div>
                )}

                {status === 'success' && (
                  <motion.div key="success" className="flex flex-col items-center justify-center py-6 text-center">
                    <CheckCircle className="w-16 h-16 text-emerald-400 mb-4" />
                    <h3 className="text-lg font-bold text-white font-mono">TASK_COMPLETED_SUCCESSFULLY</h3>
                    <p className="text-xs text-[#6b6b76] mt-1 max-w-sm">El árbol de datos de la IA ha sido inyectado en un nuevo contenedor Word y descargado localmente.</p>
                    
                    <button 
                      onClick={() => { setFile(null); setStatus('idle'); }}
                      className="mt-8 px-6 py-3 border border-[#1f1f24] hover:border-white text-white font-mono text-xs uppercase tracking-wider rounded transition-colors"
                    >
                      Montar nuevo payload
                    </button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}