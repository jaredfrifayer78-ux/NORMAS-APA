"use client";
import { useState, useRef } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';
import { motion, AnimatePresence, useScroll, useTransform } from 'framer-motion';
import { UploadCloud, FileText, CheckCircle, Loader2, ChevronLeft, ArrowRight, Layers, FileCheck2, Code, AlertTriangle, Sun, Moon } from 'lucide-react';

export default function Home() {
  // Configuración del motor de diseño: 'dark' o 'light'
  const [theme, setTheme] = useState('dark');
  const [view, setView] = useState('landing');
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState('idle'); 
  const [loadingText, setLoadingText] = useState('');

  const containerRef = useRef(null);
  const { scrollYProgress } = useScroll({ target: containerRef });
  
  const heroScale = useTransform(scrollYProgress, [0, 0.12], [1, 0.96]);
  const heroOpacity = useTransform(scrollYProgress, [0, 0.12], [1, 0]);
  const bentoY = useTransform(scrollYProgress, [0, 0.2], [40, 0]);

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
        children: [new TextRun({ text: bloque.contenido, font: "Arial", size: 24 })],
        spacing: { line: 480 },
      };

      switch(bloque.tipo) {
        case 'portada_titulo':
        case 'titulo_nivel_1':
          config.alignment = AlignmentType.CENTER;
          const titulo1 = bloque.contenido.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
          config.children[0] = new TextRun({ text: titulo1, font: "Arial", size: 24, bold: true });
          break;
        case 'titulo_nivel_2':
          config.alignment = AlignmentType.LEFT;
          const titulo2 = bloque.contenido.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
          config.children[0] = new TextRun({ text: titulo2, font: "Arial", size: 24, bold: true });
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
    a.download = "Documento_APA_Arial.docx";
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const handleProcess = async () => {
    if (!file) return;
    setStatus('loading');
    
    try {
      setLoadingText('Iniciando deserialización binaria...');
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
      const textoExtraido = result.value;
      if (!textoExtraido) throw new Error("Documento vacío o corrupto.");

      // ESCUDO DE TOKENS (PRESUPUESTO INTELIGENTE)
      // Bloqueo preventivo si el texto supera los 100,000 caracteres (~15,000-18,000 palabras)
      const MAX_CHAR_CAPACITY = 100000;
      if (textoExtraido.length > MAX_CHAR_CAPACITY) {
        throw new Error(`PAYLOAD_EXCEEDS_MAX_CAPACITY: El documento excede el límite seguro de la infraestructura gratuita (${textoExtraido.length}/${MAX_CHAR_CAPACITY} caracteres).`);
      }

      setLoadingText('Orquestando nodos con inteligencia artificial...');
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
      alert(error.message);
      setStatus('error');
    }
  };

  return (
    <div 
      ref={containerRef} 
      className={`min-h-screen font-sans antialiased overflow-x-hidden transition-colors duration-300 relative ${
        theme === 'dark' ? 'bg-[#050507] text-[#ffffff] selection:bg-indigo-500/30' : 'bg-[#f8fafc] text-[#0f172a] selection:bg-indigo-600/10'
      }`}
    >
      
      {/* Grilla técnica reactiva al tema */}
      <div className={`absolute inset-0 bg-[linear-gradient(to_right,var(--grid-color)_1px,transparent_1px),linear-gradient(to_bottom,var(--grid-color)_1px,transparent_1px)] bg-[size:5rem_5rem] [mask-image:radial-gradient(ellipse_60%_60%_at_50%_0%,#000_80%,transparent_100%)] opacity-70 pointer-events-none`}
        style={{ '--grid-color': theme === 'dark' ? '#141419' : '#e2e8f0' }}
      ></div>

      {/* Header Estilo Figma Canvas con Switcher de Diseño */}
      <nav className={`w-full h-20 border-b flex justify-between items-center px-8 fixed top-0 left-0 right-0 z-50 transition-colors backdrop-blur-xl ${
        theme === 'dark' ? 'border-[#141419] bg-[#050507]/80' : 'border-slate-200 bg-white/80 shadow-sm shadow-slate-100/50'
      }`}>
        <div className="flex items-center gap-3">
          <div className={`font-black text-xl tracking-tighter flex items-center gap-2.5 cursor-pointer ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`} onClick={() => setView('landing')}>
            <span className="w-3 h-3 bg-indigo-500 rounded-full shadow-[0_0_15px_#6366f1]"></span>
            FORMATIA <span className={`text-[10px] font-mono font-bold border px-2 py-0.5 rounded ${theme === 'dark' ? 'text-slate-500 border-[#141419] bg-[#0c0c0e]' : 'text-slate-400 border-slate-200 bg-slate-50'}`}>PRO_ENGINE</span>
          </div>
        </div>

        {/* Controles del Navbar */}
        <div className="flex items-center gap-4">
          {/* Botón Switcher de Tema */}
          <button 
            onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
            className={`p-2.5 rounded-lg border transition-all ${
              theme === 'dark' ? 'border-[#141419] bg-[#0c0c0e] text-amber-400 hover:bg-slate-800' : 'border-slate-200 bg-white text-indigo-600 hover:bg-slate-50'
            }`}
            title={theme === 'dark' ? "Cambiar a modo claro" : "Cambiar a modo oscuro"}
          >
            {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
          </button>

          <button 
            onClick={() => setView(view === 'landing' ? 'app' : 'landing')}
            className={`text-xs font-bold uppercase tracking-widest px-6 py-3 rounded-lg transition-all shadow-lg active:scale-95 ${
              theme === 'dark' ? 'bg-white text-black hover:bg-indigo-600 hover:text-white' : 'bg-slate-950 text-white hover:bg-indigo-600'
            }`}
          >
            {view === 'landing' ? 'Lanzar Workspace' : 'Cerrar Consola'}
          </button>
        </div>
      </nav>

      <AnimatePresence mode="wait">
        
        {/* =======================
            VISTA 1: LANDING PAGE REACTIVA
        ======================== */}
        {view === 'landing' && (
          <motion.div key="landing" className="pt-44 pb-40 max-w-7xl mx-auto px-8 relative z-10">
            
            {/* Hero Section */}
            <motion.div style={{ scale: heroScale, opacity: heroOpacity }} className="text-left max-w-5xl mb-32">
              <div className={`inline-flex items-center gap-2 px-3 py-1.5 mb-6 rounded-full border text-xs font-mono tracking-wider uppercase ${
                theme === 'dark' ? 'bg-indigo-950/40 border-indigo-800/40 text-indigo-400' : 'bg-indigo-50 border-indigo-100 text-indigo-600'
              }`}>
                <span>⚡ CORE_V1_OPERATIONAL</span>
              </div>
              <h1 className={`text-6xl md:text-9xl font-black tracking-tight mb-8 leading-[0.9] uppercase ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>
                Destruye el <br />
                Caos de APA.
              </h1>
              <p className={`text-xl md:text-2xl max-w-3xl mb-12 leading-relaxed font-light ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                Un motor binario puro que inyecta geometría matemática y maquetación quirúrgica a tus investigaciones. Sin tocar tus palabras, devolvemos un archivo Word impecable en Arial 12.
              </p>
              <button 
                onClick={() => setView('app')}
                className="group inline-flex items-center gap-4 bg-indigo-600 hover:bg-indigo-500 text-white font-extrabold text-lg px-10 py-5 rounded-xl transition-all shadow-[0_0_30px_rgba(99,102,241,0.3)] active:scale-95"
              >
                INGRESAR AL WORKSPACE
                <ArrowRight size={20} className="group-hover:translate-x-1 transition-transform" />
              </button>
            </motion.div>

            {/* Bento Grid */}
            <motion.div style={{ y: bentoY }} className="grid md:grid-cols-3 gap-6 max-w-6xl mx-auto text-left">
              
              {/* Tarjeta Títulos */}
              <div className={`border p-10 rounded-2xl md:col-span-2 flex flex-col justify-between transition-all group relative overflow-hidden ${
                theme === 'dark' ? 'bg-[#0c0c0e] border-[#141419] hover:border-slate-800' : 'bg-white border-slate-200 shadow-sm hover:shadow-md hover:border-slate-300'
              }`}>
                <div className="max-w-md">
                  <Layers className="text-indigo-500 mb-6 w-8 h-8" />
                  <h3 className={`text-3xl font-black mb-4 tracking-tight uppercase ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>Planchado de Títulos</h3>
                  <p className={`text-base leading-relaxed mb-8 ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                    La IA intercepta los encabezados en mayúsculas sostenidas y los normaliza dinámicamente según la jerarquía oficial de la APA, forzando la tipografía uniforme de principio a fin.
                  </p>
                </div>
                
                {/* Simulador de UI */}
                <div className={`border rounded-xl p-6 font-mono text-xs space-y-4 shadow-2xl relative ${theme === 'dark' ? 'bg-[#050507] border-[#141419]' : 'bg-slate-50 border-slate-200'}`}>
                  <span className="absolute top-2 right-3 text-[9px] text-slate-400">LIVE_COMPILER</span>
                  <div className="flex items-center gap-4 text-red-400 bg-red-950/20 p-3 rounded border border-red-900/30 opacity-60">
                    <span className="bg-red-500 text-black px-1.5 font-bold rounded-sm">RAW</span>
                    <span className="line-through tracking-wider font-bold">{"\"1. INTRODUCCIÓN GENERAL DE LA TESIS\""}</span>
                  </div>
                  <div className={`flex items-center gap-4 text-emerald-400 bg-emerald-950/20 p-3 rounded border border-emerald-900/30 shadow-[0_0_15px_rgba(16,185,129,0.1)]`}>
                    <span className="bg-emerald-400 text-black px-1.5 font-bold rounded-sm">APA</span>
                    <span className={`font-bold text-center w-full block text-sm ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>Introducción General de la Tesis</span>
                  </div>
                </div>
              </div>

              {/* Tarjeta Geometría */}
              <div className={`border p-10 rounded-2xl flex flex-col justify-between transition-all group ${
                theme === 'dark' ? 'bg-[#0c0c0e] border-[#141419] hover:border-slate-800' : 'bg-white border-slate-200 shadow-sm hover:shadow-md'
              }`}>
                <div>
                  <FileCheck2 className="text-emerald-500 mb-6 w-8 h-8" />
                  <h3 className={`text-3xl font-black mb-4 tracking-tight uppercase ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>Geometría</h3>
                  <p className={`text-base leading-relaxed ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                    Ajuste forzado de márgenes a 1 pulgada periférica constante, interlineado doble reglamentario y sangría de primera línea automatizada.
                  </p>
                </div>
                <div className={`mt-8 border h-16 rounded-xl flex items-center justify-between px-4 font-mono text-[10px] text-slate-400 relative overflow-hidden ${theme === 'dark' ? 'bg-[#050507] border-[#141419]' : 'bg-slate-50 border-slate-200'}`}>
                  <div className="absolute top-0 left-4 bottom-0 w-[1px] bg-indigo-500/50 shadow-[0_0_8px_#6366f1]"></div>
                  <span>| | | | | | | | 1.0 INCH | | | | | | | |</span>
                  <div className="absolute top-0 right-4 bottom-0 w-[1px] bg-indigo-500/50"></div>
                </div>
              </div>

              {/* Tarjeta Restricciones */}
              <div className={`border p-10 rounded-2xl md:col-span-1 flex flex-col justify-between transition-all relative overflow-hidden group ${
                theme === 'dark' ? 'bg-[#0c0c0e] border-amber-900/30 hover:border-amber-700/50' : 'bg-amber-50/40 border-amber-200 hover:bg-amber-50/80 shadow-sm'
              }`}>
                <div className="absolute top-0 right-0 w-32 h-32 bg-amber-500/5 rounded-full blur-3xl pointer-events-none"></div>
                <div>
                  <AlertTriangle className="text-amber-500 mb-6 w-8 h-8 group-hover:animate-bounce" />
                  <h3 className={`text-3xl font-black mb-4 tracking-tight uppercase ${theme === 'dark' ? 'text-white' : 'text-amber-950'}`}>Restricciones</h3>
                  <p className={`text-sm leading-relaxed ${theme === 'dark' ? 'text-slate-400' : 'text-amber-900/90'}`}>
                    Este motor está optimizado exclusivamente para **texto crudo y semántica**. Si tu archivo original contiene **tablas complejas o imágenes, estas no se mantendrán** en el mapeo de salida.
                  </p>
                </div>
                <div className={`mt-6 font-mono text-[9px] p-4 rounded-xl leading-relaxed border ${theme === 'dark' ? 'text-amber-500 bg-amber-950/20 border-amber-900/40' : 'text-amber-700 bg-amber-100/50 border-amber-200/60'}`}>
                  {"// PROTOCOLO: Procesa la estructura del texto aquí y luego reinserta tus gráficos usando copiado manual en tu Word final."}
                </div>
              </div>

              {/* Tarjeta Seguridad y Cuota */}
              <div className={`border p-10 rounded-2xl md:col-span-2 flex flex-col justify-between transition-all group ${
                theme === 'dark' ? 'bg-[#0c0c0e] border-[#141419] hover:border-slate-800' : 'bg-white border-slate-200 shadow-sm hover:shadow-md'
              }`}>
                <div>
                  <div className="flex items-center gap-2 mb-4">
                    <Code className="text-blue-500 w-5 h-5" />
                    <span className="font-mono text-xs uppercase tracking-widest text-blue-500 font-bold">INFRASTRUCTURE_METRICS</span>
                  </div>
                  <h3 className={`text-3xl font-black mb-4 tracking-tight uppercase ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>Presupuesto Seguro</h3>
                  <p className={`text-base leading-relaxed ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                    Control de tráfico local integrado. El motor valida el peso de tu payload en milisegundos para garantizar un procesamiento justo, blindando la arquitectura contra sobrecargas y asegurando privacidad absoluta en memoria RAM.
                  </p>
                </div>
                <div className="mt-8 font-mono text-[10px] text-slate-500 flex gap-4">
                  <span>[MAX_CAPACITY: 100K_CHARS]</span>
                  <span>[SANDBOX_PARSING: ACTIVE]</span>
                </div>
              </div>

            </motion.div>
          </motion.div>
        )}

        {/* =======================
            VISTA 2: EL ESPACIO DE TRABAJO (WORKSPACE)
        ======================== */}
        {view === 'app' && (
          <motion.div key="app" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="pt-32 pb-24 px-6 max-w-4xl mx-auto flex flex-col items-center justify-center min-h-screen">
            <div className="w-full max-w-2xl flex justify-start mb-4">
              <button onClick={() => { setView('landing'); setFile(null); setStatus('idle'); }} className="flex items-center gap-2 text-slate-500 hover:text-indigo-500 transition-colors font-mono text-xs uppercase tracking-widest">
                <ChevronLeft size={14} /> [ Regresar al Manifiesto ]
              </button>
            </div>

            {/* Consola de Compilación */}
            <div className={`w-full max-w-2xl border rounded-2xl p-10 md:p-14 shadow-2xl relative transition-all ${
              theme === 'dark' ? 'bg-[#0c0c0e] border-[#141419]' : 'bg-white border-slate-200'
            }`}>
              
              <div className="flex items-center gap-2 absolute top-5 left-8">
                <span className="w-2 h-2 rounded-full bg-red-500/40"></span>
                <span className="w-2 h-2 rounded-full bg-yellow-500/40"></span>
                <span className="w-2 h-2 rounded-full bg-green-500/40"></span>
              </div>

              <div className={`text-left mb-10 border-b pb-6 ${theme === 'dark' ? 'border-[#141419]' : 'border-slate-100'}`}>
                <h2 className={`text-2xl font-black uppercase tracking-tight ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>Compilador de Estructura</h2>
                <p className="text-xs text-slate-500 mt-1 font-mono">Carga el payload binario para realizar el mapeo molecular del texto.</p>
              </div>

              <AnimatePresence mode="wait">
                {(status === 'idle' || status === 'error') && (
                  <motion.div key="upload" className="flex flex-col">
                    <div className="relative w-full group">
                      <input type="file" accept=".docx" onChange={handleFileChange} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
                      <div className={`border border-dashed rounded-xl p-12 flex flex-col items-center justify-center transition-all duration-200
                        ${file ? 'border-indigo-500 bg-indigo-500/5' : theme === 'dark' ? 'border-[#141419] bg-[#050507]/40 group-hover:border-slate-700' : 'border-slate-200 bg-slate-50/50 group-hover:border-slate-300'}`}>
                        <UploadCloud className={`mb-4 ${file ? 'text-indigo-500' : 'text-slate-400'} w-12 h-12`} />
                        <h3 className={`text-sm font-bold mb-1 font-mono tracking-wider ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                          {file ? file.name.toUpperCase() : "MOUNT_PAYLOAD.DOCX"}
                        </h3>
                        <p className="text-xs text-slate-500">
                          {file ? "Buffer verificado y acoplado" : "Arrastra o selecciona el documento de origen"}
                        </p>
                      </div>
                    </div>

                    <button 
                      onClick={handleProcess}
                      disabled={!file}
                      className={`mt-8 w-full py-5 rounded-lg font-black text-xs uppercase tracking-widest transition-all font-mono border ${
                        !file 
                          ? theme === 'dark' ? 'bg-transparent text-slate-700 border-[#141419] cursor-not-allowed' : 'bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed'
                          : theme === 'dark' ? 'bg-white text-black border-white hover:bg-indigo-600 hover:text-white hover:border-indigo-600' : 'bg-slate-950 text-white border-slate-950 hover:bg-indigo-600'
                      }`}
                    >
                      Ejecutar Mapeo Sintáctico
                    </button>
                  </motion.div>
                )}

                {status === 'loading' && (
                  <motion.div key="loading" className="flex flex-col items-center justify-center py-10 text-left w-full">
                    <div className={`w-full border p-6 rounded-xl font-mono text-xs space-y-3 shadow-2xl ${theme === 'dark' ? 'bg-[#050507] border-[#141419] text-slate-500' : 'bg-slate-50 border-slate-200 text-slate-500'}`}>
                      <div className={`flex items-center gap-2 font-bold ${theme === 'dark' ? 'text-white' : 'text-slate-800'}`}>
                        <Loader2 className="w-4 h-4 text-indigo-500 animate-spin" />
                        <span>ENGINE_STATUS: EXECUTING_PIPELINE</span>
                      </div>
                      <div className="text-indigo-600 border-t border-dashed border-slate-200/60 pt-3 mt-3 animate-pulse font-bold tracking-wide">
                        &gt; {loadingText}
                      </div>
                    </div>
                  </motion.div>
                )}

                {status === 'success' && (
                  <motion.div key="success" className="flex flex-col items-center justify-center py-6 text-center">
                    <CheckCircle className="w-16 h-16 text-emerald-500 mb-6 drop-shadow-[0_0_15px_rgba(16,185,129,0.3)]" />
                    <h3 className={`text-xl font-bold font-mono uppercase tracking-wider ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>Compilación Finalizada</h3>
                    <p className="text-xs text-slate-500 mt-2 max-w-sm leading-relaxed">Los nodos semánticos estructurados han sido reinyectados en un nuevo contenedor Word binario (Arial 12) y descargados de forma inmediata.</p>
                    
                    <button 
                      onClick={() => { setFile(null); setStatus('idle'); }}
                      className={`mt-8 px-6 py-3 border font-mono text-xs uppercase tracking-widest rounded transition-all ${theme === 'dark' ? 'border-[#141419] hover:border-white text-white' : 'border-slate-200 hover:border-slate-900 text-slate-600'}`}
                    >
                      Montar nueva tarea
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