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
  
  // Animaciones de scroll suavizadas para evitar saltos bruscos en la vista
  const heroScale = useTransform(scrollYProgress, [0, 0.15], [1, 0.98]);
  const heroOpacity = useTransform(scrollYProgress, [0, 0.15], [1, 0]);
  const bentoY = useTransform(scrollYProgress, [0, 0.25], [60, 0]);

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
    <div ref={containerRef} className="min-h-screen bg-[#f8fafc] text-[#0f172a] font-sans antialiased overflow-x-hidden selection:bg-indigo-600/10 relative">
      
      {/* Grilla de diseño milimétrica en gris claro (Figma Workspace Style) */}
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#e2e8f0_1px,transparent_1px),linear-gradient(to_bottom,#e2e8f0_1px,transparent_1px)] bg-[size:3rem_3rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_0%,#000_80%,transparent_100%)] opacity-40 pointer-events-none"></div>

      {/* Navbar de Alta Precisión */}
      <nav className="w-full h-16 border-b border-slate-200 bg-white/80 backdrop-blur-md flex justify-between items-center px-6 fixed top-0 left-0 right-0 z-50 shadow-sm shadow-slate-100/50">
        <div className="flex items-center gap-6">
          <div className="font-extrabold text-base tracking-tight text-slate-900 flex items-center gap-2 cursor-pointer" onClick={() => setView('landing')}>
            <span className="w-2.5 h-2.5 bg-indigo-600 rounded-full shadow-[0_0_8px_rgba(79,70,229,0.4)]"></span>
            JareDesign Solutions <span className="text-[10px] font-mono font-normal text-slate-400 border border-slate-200 px-1.5 py-0.5 rounded bg-slate-50">v1.0.0</span>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <button 
            onClick={() => setView(view === 'landing' ? 'app' : 'landing')}
            className={`text-xs font-bold uppercase tracking-wider px-5 py-2.5 rounded transition-all border ${
              view === 'landing' 
                ? 'bg-slate-950 text-white border-slate-950 hover:bg-slate-800' 
                : 'bg-white text-slate-600 border-slate-200 hover:text-slate-900 hover:border-slate-300'
            }`}
          >
            {view === 'landing' ? 'Abrir Consola' : 'Ver Manifiesto'}
          </button>
        </div>
      </nav>

      <AnimatePresence mode="wait">
        
        {/* =======================
            VISTA 1: FIGMA LIGHT LANDING PAGE
        ======================== */}
        {view === 'landing' && (
          <motion.div 
            key="landing"
            className="pt-36 pb-32 max-w-6xl mx-auto px-6 relative z-10"
          >
            {/* Hero Section con Contraste de Alta Definición */}
            <motion.div style={{ scale: heroScale, opacity: heroOpacity }} className="text-left max-w-4xl mb-24 pt-6">
              <p className="font-mono text-xs uppercase tracking-[0.2em] text-indigo-600 mb-5 font-bold">
                {"// COMPILADOR AUTOMÁTICO DE NORMAS APA"}
              </p>
              <h1 className="text-5xl md:text-7xl font-extrabold tracking-tight text-slate-900 mb-6 leading-[1.1]">
                Estructura impecable. <br />
                Diseño automatizado.
              </h1>
              <p className="text-lg md:text-xl text-slate-600 max-w-3xl mb-10 leading-relaxed font-normal">
                No alteramos una sola palabra de tu investigación. Nuestro motor segmenta el contenido original y lo reconstruye en un archivo Word nativo bajo los estándares físicos del manual APA 7ma edición.
              </p>
              <button 
                onClick={() => setView('app')}
                className="group inline-flex items-center gap-3 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-sm uppercase tracking-wider px-7 py-4 rounded-xl transition-all shadow-md shadow-indigo-200 active:scale-95"
              >
                Inicializar Workspace 
                <ArrowRight size={16} className="group-hover:translate-x-1 transition-transform" />
              </button>
            </motion.div>

            {/* Bento Grid Asimétrico en Alta Legibilidad */}
            <motion.div style={{ y: bentoY }} className="grid md:grid-cols-3 gap-5 max-w-6xl mx-auto text-left">
              
              {/* Tarjeta 1: Títulos */}
              <div className="bg-white border border-slate-200/80 p-8 rounded-2xl md:col-span-2 flex flex-col justify-between shadow-sm hover:shadow-md hover:border-slate-300 transition-all relative overflow-hidden">
                <div>
                  <Layers className="text-indigo-600 mb-5 w-6 h-6" />
                  <h3 className="text-xl font-bold text-slate-900 mb-2 tracking-tight">Normalización de Títulos</h3>
                  <p className="text-slate-600 text-sm leading-relaxed max-w-lg">
                    El sistema detecta títulos escritos erróneamente en mayúsculas sostenidas y los transforma a formato de mayúsculas iniciales, centrados y en negrita, eliminando las correcciones manuales tediosas.
                  </p>
                </div>
                <div className="mt-8 font-mono text-[10px] text-slate-400 border-t border-slate-100 pt-4 flex gap-4">
                  <span>INPUT: TEXT_RAW</span>
                  <span>➔</span>
                  <span className="text-indigo-600 font-bold">OUTPUT: TITLE_CASE_APA</span>
                </div>
              </div>

              {/* Tarjeta 2: Geometría */}
              <div className="bg-white border border-slate-200/80 p-8 rounded-2xl flex flex-col justify-between shadow-sm hover:shadow-md hover:border-slate-300 transition-all">
                <div>
                  <FileCheck2 className="text-emerald-600 mb-5 w-6 h-6" />
                  <h3 className="text-xl font-bold text-slate-900 mb-2 tracking-tight">Geometría de Página</h3>
                  <p className="text-slate-600 text-sm leading-relaxed">
                    Márgenes exactos de 1 pulgada en los 4 flancos, interlineado doble reglamentario de 480 twips y sangrías de primera línea calibradas a 0.5 pulgadas.
                  </p>
                </div>
                <div className="mt-6 font-mono text-[10px] text-slate-400 bg-slate-50 border border-slate-100 p-2 rounded text-center">
                  [MÁRGENES: 1440 TWIPS CONSTANTE]
                </div>
              </div>

              {/* Tarjeta 3: Restricción / Honestidad del producto */}
              <div className="bg-amber-50/40 border border-amber-200/70 p-8 rounded-2xl md:col-span-1 flex flex-col justify-between shadow-sm hover:bg-amber-50/70 transition-all relative overflow-hidden">
                <div>
                  <AlertTriangle className="text-amber-600 mb-5 w-6 h-6" />
                  <h3 className="text-xl font-bold text-amber-950 mb-2 tracking-tight">Alcance del Motor</h3>
                  <p className="text-amber-900/80 text-sm leading-relaxed">
                    Esta solución procesa de forma exclusiva la **estructura de texto crudo**. Si tu documento de origen posee elementos como <span className="text-amber-800 font-bold">tablas o imágenes, estas no se exportarán</span>.
                  </p>
                </div>
                <div className="mt-6 font-mono text-[10px] text-amber-700 bg-amber-100/40 border border-amber-200/40 p-3 rounded-lg leading-relaxed">
                  {"// REGLA: Aplica el formato al texto aquí y reinserta tus gráficos manualmente en el Word resultante."}
                </div>
              </div>

              {/* Tarjeta 4: Privacidad */}
              <div className="bg-white border border-slate-200/80 p-8 rounded-2xl md:col-span-2 flex flex-col justify-between shadow-sm hover:shadow-md hover:border-slate-300 transition-all">
                <div>
                  <div className="flex items-center gap-2 mb-4">
                    <Code className="text-blue-600 w-5 h-5" />
                    <span className="font-mono text-[10px] uppercase tracking-wider text-blue-600 font-bold">Aislamiento de Datos</span>
                  </div>
                  <h3 className="text-xl font-bold text-slate-900 mb-2 tracking-tight">Privacidad por Arquitectura</h3>
                  <p className="text-slate-600 text-sm leading-relaxed">
                    La extracción del buffer binario ocurre en el entorno local de tu navegador. Tus investigaciones y propiedad intelectual jamás tocan bases de datos permanentes ni registros externos.
                  </p>
                </div>
                <div className="mt-6 font-mono text-[10px] text-slate-400">
                  [SECURE_CLIENT_PARSING: ON]
                </div>
              </div>

            </motion.div>
          </motion.div>
        )}

        {/* =======================
            VISTA 2: LA CONSOLA DE TRABAJO (WORKSPACE)
        ======================== */}
        {view === 'app' && (
          <motion.div 
            key="app"
            initial={{ opacity: 0, y: 15 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }}
            className="pt-24 pb-24 px-6 max-w-4xl mx-auto flex flex-col items-center justify-center min-h-screen"
          >
            <div className="w-full max-w-xl flex justify-start mb-4">
              <button 
                onClick={() => { setView('landing'); setFile(null); setStatus('idle'); }}
                className="flex items-center gap-1.5 text-slate-400 hover:text-slate-800 transition-colors font-mono text-[10px] uppercase tracking-wider"
              >
                <ChevronLeft size={14} /> [ Regresar al inicio ]
              </button>
            </div>

            {/* Espacio de trabajo limpio estilo interfaz Figma */}
            <div className="w-full max-w-xl bg-white border border-slate-200 rounded-2xl p-8 md:p-10 shadow-md relative">
              
              <div className="text-left mb-8 border-b border-slate-100 pb-5">
                <h2 className="text-lg font-bold text-slate-900 tracking-tight">Procesador Estructural</h2>
                <p className="text-xs text-slate-400 mt-0.5 font-mono">Monta el payload original para ejecutar el mapeo sintáctico.</p>
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
                      <div className={`border-2 border-dashed rounded-xl p-10 flex flex-col items-center justify-center transition-all duration-200
                        ${file 
                          ? 'border-indigo-500 bg-indigo-50/30' 
                          : 'border-slate-200 bg-slate-50/50 group-hover:border-slate-300 group-hover:bg-slate-50'}`}>
                        <UploadCloud className={`mb-3 ${file ? 'text-indigo-600' : 'text-slate-400'} w-8 h-8`} />
                        <h3 className="text-xs font-bold text-slate-800 mb-1 font-mono tracking-tight">
                          {file ? file.name.toUpperCase() : "SELECCIONAR_ORIGEN.DOCX"}
                        </h3>
                        <p className="text-[11px] text-slate-400 text-center">
                          {file ? "Carga verificada de datos" : "Arrastra o selecciona el archivo"}
                        </p>
                      </div>
                    </div>

                    <button 
                      onClick={handleProcess}
                      disabled={!file}
                      className={`mt-6 w-full py-4 rounded-xl font-bold text-xs uppercase tracking-widest transition-all font-mono border ${
                        !file 
                          ? 'bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed' 
                          : 'bg-slate-950 text-white border-slate-950 hover:bg-slate-800 shadow-sm'
                      }`}
                    >
                      Compilar a Formato APA
                    </button>
                  </motion.div>
                )}

                {status === 'loading' && (
                  <motion.div key="loading" className="flex flex-col items-center justify-center py-6 text-left w-full">
                    <div className="w-full bg-slate-50 border border-slate-200 p-5 rounded-xl font-mono text-xs text-slate-500 space-y-2">
                      <div className="flex items-center gap-2 font-bold text-slate-800">
                        <Loader2 className="w-3.5 h-3.5 text-indigo-600 animate-spin" />
                        <span>STATUS: PROCESSING_DATA</span>
                      </div>
                      <div className="text-indigo-600 border-t border-slate-200/60 pt-2 mt-2 animate-pulse font-medium">
                        &gt; {loadingText}
                      </div>
                    </div>
                  </motion.div>
                )}

                {status === 'success' && (
                  <motion.div key="success" className="flex flex-col items-center justify-center py-4 text-center">
                    <CheckCircle className="w-12 h-12 text-emerald-600 mb-4" />
                    <h3 className="text-base font-bold text-slate-900 font-mono">EJECUCIÓN_EXITOSA</h3>
                    <p className="text-xs text-slate-500 mt-1 max-w-xs leading-relaxed">Los nodos estructurales validados han sido inyectados en el nuevo archivo Word y descargados de forma automática.</p>
                    
                    <button 
                      onClick={() => { setFile(null); setStatus('idle'); }}
                      className="mt-6 px-5 py-2.5 border border-slate-200 hover:border-slate-400 text-slate-600 font-mono text-xs uppercase tracking-wider rounded-lg transition-colors"
                    >
                      Procesar nuevo archivo
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