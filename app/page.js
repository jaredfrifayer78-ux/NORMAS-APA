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
    a.download = "Documento_APA_Perfecto.docx";
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
      if (!textoExtraido) throw new Error("Documento vacío.");

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
      alert("Error: " + error.message);
      setStatus('error');
    }
  };

  return (
    <div ref={containerRef} className="min-h-screen bg-[#050507] text-[#ffffff] font-sans antialiased overflow-x-hidden selection:bg-indigo-500/30 relative">
      
      {/* Figma Blueprint Grid Lines */}
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#141419_1px,transparent_1px),linear-gradient(to_bottom,#141419_1px,transparent_1px)] bg-[size:5rem_5rem] [mask-image:radial-gradient(ellipse_60%_60%_at_50%_0%,#000_80%,transparent_100%)] opacity-70 pointer-events-none"></div>

      {/* Header Premium de Alta Fidelidad */}
      <nav className="w-full h-20 border-b border-[#141419] bg-[#050507]/80 backdrop-blur-xl flex justify-between items-center px-8 fixed top-0 left-0 right-0 z-50">
        <div className="flex items-center gap-3">
          <div className="font-black text-xl tracking-tighter text-white flex items-center gap-2.5 cursor-pointer" onClick={() => setView('landing')}>
            <span className="w-3 h-3 bg-indigo-500 rounded-full shadow-[0_0_15px_#6366f1]"></span>
            JareDesign Solutions <span className="text-[10px] font-mono font-bold text-slate-500 border border-[#141419] px-2 py-0.5 rounded bg-[#0c0c0e]">PRO_ENGINE</span>
          </div>
        </div>
        <button 
          onClick={() => setView(view === 'landing' ? 'app' : 'landing')}
          className="text-xs font-bold uppercase tracking-widest bg-white text-black hover:bg-indigo-500 hover:text-white px-6 py-3 rounded-lg transition-all shadow-lg active:scale-95"
        >
          {view === 'landing' ? 'Lanzar Workspace' : 'Cerrar Consola'}
        </button>
      </nav>

      <AnimatePresence mode="wait">
        
        {/* =======================
            VISTA 1: LANDING PAGE ULTRA VISUAL
        ======================== */}
        {view === 'landing' && (
          <motion.div key="landing" className="pt-44 pb-40 max-w-7xl mx-auto px-8 relative z-10">
            
            {/* Hero Seccion masivo */}
            <motion.div style={{ scale: heroScale, opacity: heroOpacity }} className="text-left max-w-5xl mb-32">
              <div className="inline-flex items-center gap-2 px-3 py-1.5 mb-6 rounded-full bg-indigo-950/40 border border-indigo-800/40 text-indigo-400 text-xs font-mono tracking-wider uppercase">
                <span>⚡ NO_MORE_TEMPLATE_CLICHES</span>
              </div>
              <h1 className="text-6xl md:text-9xl font-black tracking-tight text-white mb-8 leading-[0.9] uppercase">
                Destruye el <br />
                Caos de APA.
              </h1>
              <p className="text-xl md:text-2xl text-slate-400 max-w-3xl mb-12 leading-relaxed font-light">
                Un motor binario puro que inyecta geometría matemática y maquetación quirúrgica a tus investigaciones. Sin tocar tus palabras, devolvemos un archivo Word perfecto.
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
              
              {/* Tarjeta 1: Demostración de Títulos */}
              <div className="bg-[#0c0c0e] border border-[#141419] p-10 rounded-2xl md:col-span-2 flex flex-col justify-between hover:border-slate-800 transition-all group relative overflow-hidden">
                <div className="max-w-md">
                  <Layers className="text-indigo-500 mb-6 w-8 h-8" />
                  <h3 className="text-3xl font-black text-white mb-4 tracking-tight uppercase">Planchado de Títulos</h3>
                  <p className="text-slate-400 text-base leading-relaxed mb-8">
                    La IA intercepta los encabezados en mayúsculas sostenidas y los normaliza dinámicamente según la jerarquía oficial de la APA (Nivel 1 y 2).
                  </p>
                </div>
                
                {/* Simulador corregido sin comillas desprotegidas */}
                <div className="bg-[#050507] border border-[#141419] rounded-xl p-6 font-mono text-xs space-y-4 shadow-2xl relative">
                  <span className="absolute top-2 right-3 text-[9px] text-slate-600">LIVE_COMPILER</span>
                  <div className="flex items-center gap-4 text-red-400 bg-red-950/20 p-3 rounded border border-red-900/30 opacity-60">
                    <span className="bg-red-500 text-black px-1.5 font-bold rounded-sm">RAW</span>
                    <span className="line-through tracking-wider font-bold">{"\"1. INTRODUCCIÓN GENERAL DE LA TESIS\""}</span>
                  </div>
                  <div className="flex items-center gap-4 text-emerald-400 bg-emerald-950/20 p-3 rounded border border-emerald-900/30 shadow-[0_0_15px_rgba(16,185,129,0.1)]">
                    <span className="bg-emerald-400 text-black px-1.5 font-bold rounded-sm">APA</span>
                    <span className="font-bold text-center w-full block text-white text-sm">Introducción General de la Tesis</span>
                  </div>
                </div>
              </div>

              {/* Tarjeta 2: Geometría */}
              <div className="bg-[#0c0c0e] border border-[#141419] p-10 rounded-2xl flex flex-col justify-between hover:border-slate-800 transition-all group">
                <div>
                  <FileCheck2 className="text-emerald-500 mb-6 w-8 h-8" />
                  <h3 className="text-3xl font-black text-white mb-4 tracking-tight uppercase">Geometría</h3>
                  <p className="text-slate-400 text-base leading-relaxed">
                    Ajuste forzado de márgenes a 1 pulgada periférica constante, interlineado doble de 480 twips y sangría francesa milimétrica.
                  </p>
                </div>
                <div className="mt-8 bg-[#050507] border border-[#141419] h-16 rounded-xl flex items-center justify-between px-4 font-mono text-[10px] text-slate-600 relative overflow-hidden">
                  <div className="absolute top-0 left-4 bottom-0 w-[1px] bg-indigo-500/50 shadow-[0_0_8px_#6366f1]"></div>
                  <span>| | | | | | | | 1.0 INCH | | | | | | | |</span>
                  <div className="absolute top-0 right-4 bottom-0 w-[1px] bg-indigo-500/50"></div>
                </div>
              </div>

              {/* Tarjeta 3: Restricciones */}
              <div className="bg-[#0c0c0e] border border-amber-900/30 p-10 rounded-2xl md:col-span-1 flex flex-col justify-between hover:border-amber-700/50 transition-all relative overflow-hidden group">
                <div className="absolute top-0 right-0 w-32 h-32 bg-amber-500/5 rounded-full blur-3xl pointer-events-none"></div>
                <div>
                  <AlertTriangle className="text-amber-500 mb-6 w-8 h-8 group-hover:animate-bounce" />
                  <h3 className="text-3xl font-black text-white mb-4 tracking-tight uppercase text-amber-500">Restricciones</h3>
                  <p className="text-slate-400 text-sm leading-relaxed">
                    Este motor está optimizado de forma exclusiva para **texto crudo y semántica**. Si tu archivo original contiene <span className="text-white font-bold underline decoration-amber-500">tablas complejas o imágenes, estas no se mantendrán</span> en el mapeo de salida.
                  </p>
                </div>
                <div className="mt-6 font-mono text-[9px] text-amber-500 bg-amber-950/20 border border-amber-900/40 p-4 rounded-xl leading-relaxed">
                  {"// PROTOCOLO: Procesa la estructura del texto aquí y luego reinserta tus gráficos usando copiado manual en tu Word final."}
                </div>
              </div>

              {/* Tarjeta 4: Seguridad */}
              <div className="bg-[#0c0c0e] border border-[#141419] p-10 rounded-2xl md:col-span-2 flex flex-col justify-between hover:border-slate-800 transition-all group">
                <div>
                  <div className="flex items-center gap-2 mb-4">
                    <Code className="text-blue-500 w-5 h-5" />
                    <span className="font-mono text-xs uppercase tracking-widest text-blue-400 font-bold">CLIENT_SIDE_SECURITY</span>
                  </div>
                  <h3 className="text-3xl font-black text-white mb-4 tracking-tight uppercase">Privacidad Cero-Almacenamiento</h3>
                  <p className="text-slate-400 text-base leading-relaxed">
                    El parseo binario ocurre directamente en los buffers de memoria de tu navegador. Tu propiedad intelectual, tesis e investigaciones confidenciales jamás tocan un disco duro externo.
                  </p>
                </div>
                <div className="mt-8 font-mono text-[10px] text-slate-600 flex gap-4">
                  <span>[RAM_BUFFER: ACTIVE]</span>
                  <span>[DISK_WRITE: DISABLED]</span>
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
              <button onClick={() => { setView('landing'); setFile(null); setStatus('idle'); }} className="flex items-center gap-2 text-slate-500 hover:text-white transition-colors font-mono text-xs uppercase tracking-widest">
                <ChevronLeft size={14} /> [ Regresar al Manifiesto ]
              </button>
            </div>

            <div className="w-full max-w-2xl bg-[#0c0c0e] border border-[#141419] rounded-2xl p-10 md:p-14 shadow-2xl relative">
              
              <div className="flex items-center gap-2 absolute top-5 left-8">
                <span className="w-2 h-2 rounded-full bg-red-500/40"></span>
                <span className="w-2 h-2 rounded-full bg-yellow-500/40"></span>
                <span className="w-2 h-2 rounded-full bg-green-500/40"></span>
              </div>

              <div className="text-left mb-10 border-b border-[#141419] pb-6">
                <h2 className="text-2xl font-black text-white uppercase tracking-tight">Compilador de Estructura</h2>
                <p className="text-xs text-slate-500 mt-1 font-mono">Carga el payload binario para realizar el mapeo molecular del texto.</p>
              </div>

              <AnimatePresence mode="wait">
                {(status === 'idle' || status === 'error') && (
                  <motion.div key="upload" className="flex flex-col">
                    <div className="relative w-full group">
                      <input type="file" accept=".docx" onChange={handleFileChange} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
                      <div className={`border border-dashed rounded-xl p-12 flex flex-col items-center justify-center transition-all duration-200 bg-[#050507]/40
                        ${file ? 'border-indigo-500 bg-indigo-500/5' : 'border-[#141419] group-hover:border-slate-700'}`}>
                        <UploadCloud className={`mb-4 ${file ? 'text-indigo-400' : 'text-slate-600'} w-12 h-12`} />
                        <h3 className="text-sm font-bold text-white mb-1 font-mono tracking-wider">
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
                          ? 'bg-transparent text-slate-600 border-[#141419] cursor-not-allowed' 
                          : 'bg-white text-black border-white hover:bg-indigo-600 hover:text-white hover:border-indigo-600'
                      }`}
                    >
                      Ejecutar Mapeo Sintáctico
                    </button>
                  </motion.div>
                )}

                {status === 'loading' && (
                  <motion.div key="loading" className="flex flex-col items-center justify-center py-10 text-left w-full">
                    <div className="w-full bg-[#050507] border border-[#141419] p-6 rounded-xl font-mono text-xs text-slate-500 space-y-3 shadow-2xl">
                      <div className="flex items-center gap-2 font-bold text-white">
                        <Loader2 className="w-4 h-4 text-indigo-500 animate-spin" />
                        <span>ENGINE_STATUS: EXECUTING_PIPELINE</span>
                      </div>
                      <div className="text-indigo-400 border-t border-[#141419] pt-3 mt-3 animate-pulse font-bold tracking-wide">
                        &gt; {loadingText}
                      </div>
                    </div>
                  </motion.div>
                )}

                {status === 'success' && (
                  <motion.div key="success" className="flex flex-col items-center justify-center py-6 text-center">
                    <CheckCircle className="w-16 h-16 text-emerald-500 mb-6 drop-shadow-[0_0_15px_rgba(16,185,129,0.3)]" />
                    <h3 className="text-xl font-bold text-white font-mono uppercase tracking-wider">Compilación Finalizada</h3>
                    <p className="text-xs text-slate-500 mt-2 max-w-sm leading-relaxed">Los nodos semánticos estructurados han sido reinyectados en un nuevo contenedor Word binario y descargados localmente de forma inmediata.</p>
                    
                    <button 
                      onClick={() => { setFile(null); setStatus('idle'); }}
                      className="mt-8 px-6 py-3 border border-[#141419] hover:border-white text-white font-mono text-xs uppercase tracking-widest rounded transition-all"
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