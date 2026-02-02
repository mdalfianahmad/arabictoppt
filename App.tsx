
import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import { Upload, FileSpreadsheet, Layout, PlayCircle, CheckCircle, AlertCircle, ArrowLeft, Download, RotateCcw, ChevronRight, Layers, Settings2, Maximize, Move, Ruler, List, Info, HelpCircle, BookOpen, MousePointer2, FileText, Presentation } from 'lucide-react';
import { AppStep, AppState, ColumnMapping, RowSelectionMode, IndexPosition, ColumnStyle, AspectRatio, ThemeMode } from './types';
import { parseExcelFile, getDefaultMapping } from './services/excelService';
import { generatePptx } from './services/pptxService';

const App: React.FC = () => {
  const [state, setState] = useState<AppState>({
    step: AppStep.INTRO,
    file: null,
    data: [],
    columns: [],
    mapping: { 
      box1: { column: '', fontSize: 30, x: 5, y: 10, w: 90, h: 15 },
      box2: { column: '', fontSize: 30, x: 5, y: 30, w: 90, h: 15 },
      box3: { column: '', fontSize: 26, x: 5, y: 55, w: 90, h: 15 },
      box4: { column: '', fontSize: 20, x: 5, y: 75, w: 43, h: 10 },
      box5: { column: '', fontSize: 20, x: 52, y: 75, w: 43, h: 10 }
    },
    isGenerating: false,
    error: null,
    selectionMode: 'all',
    rangeStart: 1,
    rangeEnd: 10,
    specificRows: '',
    includeIndex: true,
    indexPosition: 'content-bottom-right',
    indexCol1: '',
    indexCol2: '',
    indexFormat: '{1}:{2}',
    aspectRatio: '16:9',
    slideWidth: 1920,
    slideHeight: 1080,
    themeMode: 'light'
  });

  const [hoveredField, setHoveredField] = useState<string | null>(null);
  const [activeSettingsField, setActiveSettingsField] = useState<string | null>(null);
  const [selectedDesignerBox, setSelectedDesignerBox] = useState<string | null>(null);
  const [isDraggingFile, setIsDraggingFile] = useState(false);
  
  const previewRef = useRef<HTMLDivElement>(null);
  const previewSlideRef = useRef<HTMLDivElement>(null);
  const [previewScale, setPreviewScale] = useState(1);
  const [dragAction, setDragAction] = useState<{ key: keyof ColumnMapping, type: 'move' | 'resize', startX: number, startY: number, initialX: number, initialY: number, initialW: number, initialH: number } | null>(null);

  const processFile = async (file: File) => {
    const isSupported = file.name.endsWith('.xlsx') || file.name.endsWith('.csv');
    if (!isSupported) {
      setState(prev => ({ ...prev, error: "Only .xlsx and .csv files are supported." }));
      return;
    }
    try {
      const { data, columns } = await parseExcelFile(file);
      setState(prev => ({
        ...prev, 
        file, 
        data, 
        columns, 
        error: null, 
        step: AppStep.MAPPING, 
        rangeEnd: data.length,
        mapping: getDefaultMapping(columns),
        indexCol1: columns[0] || '',
        indexCol2: columns[1] || ''
      }));
    } catch (err: any) {
      setState(prev => ({ ...prev, error: err.message }));
    }
  };

  const handleMouseMove = useCallback((e: MouseEvent) => {
    if (!dragAction || !previewRef.current) return;
    const rect = previewRef.current.getBoundingClientRect();
    const deltaX = ((e.clientX - dragAction.startX) / rect.width) * 100;
    const deltaY = ((e.clientY - dragAction.startY) / rect.height) * 100;

    setState(prev => {
      const current = prev.mapping[dragAction.key];
      const nextMapping = { ...prev.mapping };
      
      if (dragAction.type === 'move') {
        nextMapping[dragAction.key] = {
          ...current,
          x: Math.round(Math.max(0, Math.min(100 - current.w, dragAction.initialX + deltaX))),
          y: Math.round(Math.max(0, Math.min(100 - current.h, dragAction.initialY + deltaY)))
        };
      } else {
        nextMapping[dragAction.key] = {
          ...current,
          w: Math.round(Math.max(5, Math.min(100 - current.x, dragAction.initialW + deltaX))),
          h: Math.round(Math.max(5, Math.min(100 - current.y, dragAction.initialH + deltaY)))
        };
      }
      return { ...prev, mapping: nextMapping };
    });
  }, [dragAction]);

  const handleMouseUp = useCallback(() => setDragAction(null), []);

  useEffect(() => {
    if (dragAction) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [dragAction, handleMouseMove, handleMouseUp]);

  useEffect(() => {
    const el = previewSlideRef.current;
    if (!el) return;

    const update = () => {
      const rect = el.getBoundingClientRect();
      const denom = state.slideWidth || 1;
      setPreviewScale(rect.width / denom);
    };

    update();

    const ro = new ResizeObserver(update);
    ro.observe(el);
    window.addEventListener('resize', update);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', update);
    };
  }, [state.slideWidth, state.step]);

  const ptToPreviewPx = useCallback((pt: number) => {
    // PPT points -> CSS px at 96dpi, then scaled to the preview's rendered width.
    return pt * (96 / 72) * previewScale;
  }, [previewScale]);

  const startInteraction = (e: React.MouseEvent, key: keyof ColumnMapping, type: 'move' | 'resize') => {
    e.stopPropagation();
    setSelectedDesignerBox(key);
    const style = state.mapping[key];
    setDragAction({
      key, type, startX: e.clientX, startY: e.clientY,
      initialX: style.x, initialY: style.y, initialW: style.w, initialH: style.h
    });
  };

  const updateMappingField = (key: keyof ColumnMapping, field: keyof ColumnStyle, value: any) => {
    setState(prev => ({
      ...prev,
      mapping: { ...prev.mapping, [key]: { ...prev.mapping[key], [field]: value } }
    } as AppState));
  };

  const resetLayout = () => {
    if (state.columns.length > 0) {
      setState(prev => ({ ...prev, mapping: getDefaultMapping(prev.columns) }));
    }
  };

  const setDimensions = (w: number, h: number) => {
    let ratio: AspectRatio = 'Custom';
    if (w === 1920 && h === 1080) ratio = '16:9';
    else if (w === 1440 && h === 1080) ratio = '4:3';
    
    setState(prev => ({ ...prev, slideWidth: w, slideHeight: h, aspectRatio: ratio }));
  };

  const filteredData = useMemo(() => {
    const { selectionMode, data, rangeStart, rangeEnd, specificRows } = state;
    if (selectionMode === 'all') return data;
    if (selectionMode === 'range') return data.slice(Math.max(1, rangeStart) - 1, Math.min(data.length, rangeEnd));
    if (selectionMode === 'specific') {
      const idx = specificRows.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n) && n >= 1 && n <= data.length).map(n => n - 1);
      return idx.map(i => data[i]);
    }
    return data;
  }, [state]);

  const handleGenerate = async () => {
    if (!state.file || filteredData.length === 0) return;
    setState(prev => ({ ...prev, isGenerating: true }));
    try {
      await generatePptx(
        filteredData, state.mapping, state.file.name, state.columns, 
        state.includeIndex, state.indexPosition, state.indexCol1, state.indexCol2, state.indexFormat,
        state.aspectRatio, state.slideWidth, state.slideHeight, state.themeMode
      );
      setState(prev => ({ ...prev, step: AppStep.DOWNLOAD, isGenerating: false }));
    } catch (err: any) {
      setState(prev => ({ ...prev, error: err.message, isGenerating: false }));
    }
  };

  const PrecisionInput = ({ label, value, min, max, onChange }: any) => (
    <div className="space-y-1.5">
      <div className="flex justify-between items-center px-1">
        <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">{label}</label>
      </div>
      <div className="flex items-center gap-3">
        <input 
          type="range" min={min} max={max} value={value} 
          onChange={(e) => onChange(parseInt(e.target.value))}
          className="flex-1 h-1 accent-black bg-zinc-100 rounded-lg appearance-none cursor-pointer"
        />
        <input 
          type="number" value={value} min={min} max={max}
          className="w-16 bg-zinc-800 text-white text-[11px] font-bold border-none rounded-lg px-2 py-1.5 text-center focus:ring-2 focus:ring-white/20 outline-none shadow-lg"
          onChange={(e) => onChange(parseInt(e.target.value) || 0)}
        />
      </div>
    </div>
  );

  const getPreviewIndex = () => {
    if (!state.includeIndex || filteredData.length === 0) return '';
    const val1 = state.indexCol1 ? String(filteredData[0][state.indexCol1] || '') : '';
    const val2 = state.indexCol2 ? String(filteredData[0][state.indexCol2] || '') : '';
    let fmt = state.indexFormat || '{1}:{2}';
    return fmt.replace('{1}', val1).replace('{2}', val2);
  };

  const themeBg = state.themeMode === 'dark' ? '#000000' : '#FFFFFF';
  const themeText = state.themeMode === 'dark' ? '#FFFFFF' : '#000000';

  return (
    <div className="min-h-screen bg-[#F8F9FA] flex flex-col items-center py-10 px-4">
      {/* Brand Header */}
      <div className="w-full max-w-6xl flex items-center justify-between mb-8">
        <div className="flex items-center gap-3 group cursor-pointer" onClick={() => window.location.reload()}>
          <div className="bg-black p-2.5 rounded-2xl shadow-lg transform group-hover:rotate-12 transition-transform">
            <Layout size={24} className="text-white" />
          </div>
          <span className="text-2xl font-black tracking-tighter">ARABICSLIDES</span>
        </div>
        <div className="flex gap-4">
          {state.step !== AppStep.INTRO && (
            <button onClick={() => window.location.reload()} className="flex items-center gap-2 text-zinc-400 hover:text-black font-bold text-xs tracking-widest uppercase transition-colors">
              <RotateCcw size={14} /> NEW PROJECT
            </button>
          )}
        </div>
      </div>

      <div className="w-full max-w-6xl bg-white border border-zinc-100 rounded-[3rem] shadow-2xl overflow-hidden flex flex-col">
        {/* Navigation Tabs */}
        <div className="flex border-b border-zinc-50 bg-zinc-50/10">
          {[
            { id: AppStep.INTRO, label: 'Home', icon: Info },
            { id: AppStep.UPLOAD, label: 'Upload', icon: Upload },
            { id: AppStep.MAPPING, label: 'Designer', icon: FileSpreadsheet },
            { id: AppStep.PREVIEW, label: 'Review', icon: PlayCircle },
            { id: AppStep.DOWNLOAD, label: 'Export', icon: CheckCircle }
          ].map(tab => (
            <div key={tab.id} className={`flex-1 flex items-center justify-center gap-2.5 py-6 px-6 border-b-2 transition-all duration-500 ${state.step === tab.id ? 'border-black text-black bg-white' : 'border-transparent text-zinc-300'}`}>
              <tab.icon size={18} />
              <span className="text-xs font-black uppercase tracking-widest hidden md:inline">{tab.label}</span>
            </div>
          ))}
        </div>

        <div className="p-8 md:p-14">
          {state.error && (
            <div className="mb-10 bg-red-50 border border-red-100 text-red-900 px-6 py-4 rounded-2xl flex items-center gap-3 animate-in fade-in slide-in-from-top-4">
              <AlertCircle size={20} className="text-red-400" />
              <p className="font-bold text-xs uppercase tracking-tight">{state.error}</p>
            </div>
          )}

          {state.step === AppStep.INTRO && (
            <div className="animate-in fade-in py-10 flex flex-col items-center">
              <div className="text-center mb-16 max-w-3xl">
                <h2 className="text-4xl font-black text-zinc-900 mb-6 tracking-tight">Professional Arabic Decks in Seconds</h2>
                <p className="text-zinc-500 text-lg font-medium leading-relaxed">
                  Turn your Arabic vocabulary lists into polished PowerPoint presentations automatically. No manual formatting, no copy-pasting required.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-12 w-full max-w-5xl mb-20">
                <div className="space-y-8">
                  <div className="flex gap-6 items-start">
                    <div className="bg-black text-white w-10 h-10 rounded-2xl flex items-center justify-center shrink-0 font-black text-sm">1</div>
                    <div className="space-y-3">
                      <h3 className="text-xl font-black uppercase tracking-tight">Prepare your Excel</h3>
                      <p className="text-zinc-400 text-sm leading-relaxed font-medium">Your file should have clear columns. We recommend including Arabic, Transliteration, and Translation.</p>
                      
                      <div className="bg-zinc-50 border border-zinc-100 rounded-2xl p-4 shadow-inner">
                        <div className="grid grid-cols-3 gap-2 mb-2">
                          {['Arabic', 'Translit', 'English'].map(h => <div key={h} className="text-[9px] font-black uppercase tracking-widest text-zinc-400 border-b border-zinc-100 pb-1">{h}</div>)}
                        </div>
                        <div className="grid grid-cols-3 gap-2 opacity-60">
                          {['كتاب', 'Kitāb', 'Book'].map(c => <div key={c} className="text-[10px] font-bold text-zinc-900 truncate">{c}</div>)}
                          {['قلم', 'Qalam', 'Pen'].map(c => <div key={c} className="text-[10px] font-bold text-zinc-900 truncate">{c}</div>)}
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="flex gap-6 items-start">
                    <div className="bg-black text-white w-10 h-10 rounded-2xl flex items-center justify-center shrink-0 font-black text-sm">2</div>
                    <div className="space-y-3">
                      <h3 className="text-xl font-black uppercase tracking-tight">Upload & Design</h3>
                      <p className="text-zinc-400 text-sm leading-relaxed font-medium">Upload your file and drag boxes around the Designer to create your perfect slide layout.</p>
                    </div>
                  </div>
                </div>

                <div className="space-y-8">
                  <div className="flex gap-6 items-start">
                    <div className="bg-black text-white w-10 h-10 rounded-2xl flex items-center justify-center shrink-0 font-black text-sm">3</div>
                    <div className="space-y-3">
                      <h3 className="text-xl font-black uppercase tracking-tight">Review & Export</h3>
                      <p className="text-zinc-400 text-sm leading-relaxed font-medium">Verify your slides in the Review tab, then export everything to a professional PPTX file.</p>
                      
                      <div className="bg-black rounded-2xl p-1.5 shadow-2xl aspect-video relative overflow-hidden ring-4 ring-zinc-50">
                        <div className="absolute inset-0 bg-white m-1 rounded-xl flex flex-col items-center justify-center space-y-2">
                          <div className="text-2xl font-black arabic-font">كتاب</div>
                          <div className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Kitāb</div>
                          <div className="text-sm font-bold text-zinc-600">Book</div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <button 
                onClick={() => setState(prev => ({ ...prev, step: AppStep.UPLOAD }))}
                className="bg-black text-white font-black py-7 px-16 rounded-[2.5rem] text-xl shadow-2xl hover:bg-zinc-800 transition-all transform hover:scale-105 flex items-center gap-4 group"
              >
                GET STARTED <ChevronRight className="group-hover:translate-x-1 transition-transform" />
              </button>
            </div>
          )}

          {state.step === AppStep.UPLOAD && (
            <div className="flex flex-col items-center text-center animate-in fade-in py-20">
              <h2 className="text-4xl font-black text-zinc-900 mb-4 tracking-tight">Upload Source</h2>
              <p className="text-zinc-400 mb-14 max-w-sm font-medium leading-relaxed">Drop your Excel or CSV file here to begin the automation process.</p>
              
              <label 
                onDrop={(e) => { e.preventDefault(); setIsDraggingFile(false); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); }}
                onDragOver={(e) => { e.preventDefault(); setIsDraggingFile(true); }}
                onDragLeave={() => setIsDraggingFile(false)}
                className="w-full max-w-xl cursor-pointer group"
              >
                <div className={`border-2 border-dashed rounded-[3rem] p-24 flex flex-col items-center transition-all duration-700 ${isDraggingFile ? 'border-black bg-zinc-50 scale-105 shadow-2xl' : 'border-zinc-100 bg-zinc-50/10 hover:border-black hover:bg-white'}`}>
                  <div className="bg-white p-6 rounded-3xl shadow-xl mb-8 group-hover:scale-110 transition-transform">
                    <FileSpreadsheet className={`${isDraggingFile ? 'text-black' : 'text-zinc-200'}`} size={64} />
                  </div>
                  <p className="text-zinc-900 font-black text-xl">Drop XLSX or CSV</p>
                  <p className="text-zinc-400 text-[10px] font-black tracking-[0.3em] uppercase mt-4">Drag & drop to begin</p>
                </div>
                <input type="file" className="hidden" accept=".xlsx,.csv" onChange={(e) => e.target.files?.[0] && processFile(e.target.files[0])} />
              </label>
              
              <button onClick={() => setState(prev => ({ ...prev, step: AppStep.INTRO }))} className="mt-12 text-zinc-400 hover:text-black font-black text-xs uppercase tracking-widest flex items-center gap-2 transition-colors">
                <ArrowLeft size={14} /> BACK TO GUIDE
              </button>
            </div>
          )}

          {state.step === AppStep.MAPPING && (
            <div className="animate-in fade-in flex flex-col gap-12">
              {/* 1. Box Configuration Section */}
              <section className="space-y-8">
                <div className="flex items-center justify-between">
                  <h2 className="text-2xl font-black text-zinc-900 flex items-center gap-3">
                    <Settings2 size={24} /> 1. Box Configuration
                    <div className="group relative">
                      <HelpCircle size={16} className="text-zinc-300 cursor-help" />
                      <div className="absolute left-0 top-full mt-2 hidden group-hover:block bg-black text-white text-[10px] font-bold py-2 px-3 rounded-xl w-48 shadow-2xl z-50">
                        Map Excel columns to specific boxes on your slides. Click the gear icon to adjust position and size manually.
                      </div>
                    </div>
                  </h2>
                  <button onClick={resetLayout} className="text-[10px] font-black text-zinc-400 hover:text-black uppercase tracking-widest flex items-center gap-1.5 transition-colors">
                    <RotateCcw size={12} /> RESTORE DEFAULTS
                  </button>
                </div>
                
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {(Object.entries(state.mapping) as [keyof ColumnMapping, ColumnStyle][]).map(([key, style], i) => (
                    <div 
                      key={key} 
                      className={`p-6 rounded-[2rem] border transition-all duration-300 ${hoveredField === key || selectedDesignerBox === key ? 'border-black ring-1 ring-black/5 bg-zinc-50/50 shadow-sm' : 'border-zinc-50 bg-zinc-50/20'}`} 
                      onMouseEnter={() => setHoveredField(key)} 
                      onMouseLeave={() => setHoveredField(null)}
                    >
                      <div className="flex items-center gap-4">
                        <div className="flex-1">
                          <label className="text-[10px] font-black text-zinc-300 uppercase tracking-widest mb-2 block">Box {i + 1}</label>
                          <div className="relative">
                            <select 
                              value={style.column} 
                              onChange={(e) => updateMappingField(key as any, 'column', e.target.value)}
                              className="w-full bg-white border border-zinc-100 rounded-2xl px-5 py-3.5 text-sm font-bold appearance-none shadow-sm focus:outline-none focus:ring-1 focus:ring-black"
                            >
                              <option value="">— LEAVE EMPTY —</option>
                              {state.columns.map(col => <option key={col} value={col}>{col}</option>)}
                            </select>
                            <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-zinc-300"><ChevronRight size={16} className="rotate-90" /></div>
                          </div>
                        </div>
                        <button 
                          onClick={() => setActiveSettingsField(activeSettingsField === key ? null : key)}
                          className={`p-4 rounded-2xl transition-all self-end shadow-sm ${activeSettingsField === key ? 'bg-black text-white' : 'text-zinc-300 hover:text-black bg-white border border-zinc-100'}`}
                        >
                          <Settings2 size={18} />
                        </button>
                      </div>
                      {activeSettingsField === key && (
                        <div className="mt-8 pt-8 border-t border-zinc-100 grid grid-cols-1 gap-x-6 gap-y-4 animate-in slide-in-from-top-4 duration-500">
                          <PrecisionInput label="Text Size (PT)" value={style.fontSize} min={8} max={120} onChange={(v: any) => updateMappingField(key as any, 'fontSize', v)} />
                          <PrecisionInput label="Pos X (%)" value={style.x} min={0} max={100} onChange={(v: any) => updateMappingField(key as any, 'x', v)} />
                          <PrecisionInput label="Pos Y (%)" value={style.y} min={0} max={100} onChange={(v: any) => updateMappingField(key as any, 'y', v)} />
                          <PrecisionInput label="Width (%)" value={style.w} min={5} max={100} onChange={(v: any) => updateMappingField(key as any, 'w', v)} />
                          <PrecisionInput label="Height (%)" value={style.h} min={5} max={100} onChange={(v: any) => updateMappingField(key as any, 'h', v)} />
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </section>

              {/* 2. Slide Dimensions Section */}
              <section className="space-y-8 border-t border-zinc-50 pt-12">
                <h2 className="text-2xl font-black text-zinc-900 flex items-center gap-3">
                  <Ruler size={24} /> 2. Slide Dimensions & Index
                  <div className="group relative">
                    <HelpCircle size={16} className="text-zinc-300 cursor-help" />
                    <div className="absolute left-0 top-full mt-2 hidden group-hover:block bg-black text-white text-[10px] font-bold py-2 px-3 rounded-xl w-48 shadow-2xl z-50">
                      Set output resolution and index formatting.
                    </div>
                  </div>
                </h2>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                  <div className="bg-zinc-50/30 p-8 rounded-[2.5rem] border border-zinc-50">
                    <div className="flex items-center justify-between mb-6">
                       <label className="text-[10px] font-black text-zinc-300 uppercase tracking-widest block">Size (Aspect Ratio: {state.aspectRatio})</label>
                    </div>
                    <div className="flex gap-3">
                      <button onClick={() => setDimensions(1920, 1080)} className={`flex-1 py-4 rounded-2xl font-black text-xs border transition-all ${state.slideWidth === 1920 && state.slideHeight === 1080 ? 'bg-black text-white border-black shadow-lg' : 'bg-white text-zinc-400 border-zinc-100'}`}>16:9 WIDE</button>
                      <button onClick={() => setDimensions(1440, 1080)} className={`flex-1 py-4 rounded-2xl font-black text-xs border transition-all ${state.slideWidth === 1440 && state.slideHeight === 1080 ? 'bg-black text-white border-black shadow-lg' : 'bg-white text-zinc-400 border-zinc-100'}`}>4:3 STANDARD</button>
                    </div>
                    <div className="mt-8 grid grid-cols-2 gap-4">
                       <div className="space-y-2">
                          <label className="text-[9px] font-black text-zinc-300 uppercase tracking-widest">Width (PX)</label>
                          <input 
                            type="number" value={state.slideWidth} 
                            onChange={(e) => setDimensions(parseInt(e.target.value) || 0, state.slideHeight)}
                            className="w-full bg-zinc-800 text-white font-bold py-2.5 px-3 rounded-xl focus:ring-2 focus:ring-white/20 outline-none shadow-md"
                          />
                       </div>
                       <div className="space-y-2">
                          <label className="text-[9px] font-black text-zinc-300 uppercase tracking-widest">Height (PX)</label>
                          <input 
                            type="number" value={state.slideHeight} 
                            onChange={(e) => setDimensions(state.slideWidth, parseInt(e.target.value) || 0)}
                            className="w-full bg-zinc-800 text-white font-bold py-2.5 px-3 rounded-xl focus:ring-2 focus:ring-white/20 outline-none shadow-md"
                          />
                       </div>
                    </div>
                  </div>

                  <div className="bg-zinc-50/30 p-8 rounded-[2.5rem] border border-zinc-50 flex flex-col justify-start">
                    <div className="flex items-center gap-3 mb-6">
                      <label className="text-[10px] font-black text-zinc-300 uppercase tracking-widest block">Index Placement</label>
                    </div>

                    <div className="mb-8">
                      <label className="text-[9px] font-black text-zinc-300 uppercase tracking-widest block mb-3">Theme (PPTX Output)</label>
                      <div className="flex gap-2">
                        {(['light', 'dark'] as ThemeMode[]).map(mode => (
                          <button
                            key={mode}
                            onClick={() => setState(prev => ({ ...prev, themeMode: mode }))}
                            className={`flex-1 py-3 rounded-xl text-[10px] font-black uppercase tracking-widest border transition-all ${
                              state.themeMode === mode
                                ? 'bg-black text-white border-black shadow-md'
                                : 'bg-white text-zinc-300 border-zinc-100 hover:text-black hover:border-black'
                            }`}
                          >
                            {mode}
                          </button>
                        ))}
                      </div>
                      <p className="mt-3 text-[10px] font-bold text-zinc-400 leading-relaxed">
                        Light mode exports with white background + black text. Dark mode exports with black background + white text.
                      </p>
                    </div>

                    <div className="flex items-center gap-4 mb-6 group cursor-pointer" onClick={() => setState(prev => ({ ...prev, includeIndex: !prev.includeIndex }))}>
                      <div className={`w-12 h-6 rounded-full transition-colors relative ${state.includeIndex ? 'bg-black' : 'bg-zinc-200'}`}>
                        <div className={`absolute top-1 w-4 h-4 bg-white rounded-full transition-transform ${state.includeIndex ? 'translate-x-7' : 'translate-x-1'}`} />
                      </div>
                      <span className="text-xs font-black uppercase tracking-widest">Enable Index</span>
                    </div>

                    {state.includeIndex && (
                      <div className="space-y-4 animate-in fade-in zoom-in-95 duration-500">
                        <div className="grid grid-cols-2 gap-2">
                          <select 
                            value={state.indexCol1} 
                            onChange={(e) => setState(prev => ({ ...prev, indexCol1: e.target.value }))}
                            className="w-full bg-white border border-zinc-100 rounded-xl px-3 py-2 text-[11px] font-bold appearance-none shadow-sm focus:ring-1 focus:ring-black outline-none"
                          >
                            <option value="">Col 1</option>
                            {state.columns.map(col => <option key={col} value={col}>{col}</option>)}
                          </select>
                          <select 
                            value={state.indexCol2} 
                            onChange={(e) => setState(prev => ({ ...prev, indexCol2: e.target.value }))}
                            className="w-full bg-white border border-zinc-100 rounded-xl px-3 py-2 text-[11px] font-bold appearance-none shadow-sm focus:ring-1 focus:ring-black outline-none"
                          >
                            <option value="">Col 2</option>
                            {state.columns.map(col => <option key={col} value={col}>{col}</option>)}
                          </select>
                        </div>

                        <div className="relative group/fmt">
                          <div className="flex items-center justify-between mb-1.5 px-1">
                            <label className="text-[9px] font-black text-zinc-300 uppercase tracking-widest">Index Format</label>
                            <HelpCircle size={12} className="text-zinc-300 cursor-help" />
                            <div className="absolute right-0 bottom-full mb-2 hidden group-hover/fmt:block bg-zinc-800 text-white text-[9px] font-bold py-2 px-3 rounded-lg w-48 shadow-2xl z-50">
                              Use <span className="text-zinc-400">{"{1}"}</span> for column 1 and <span className="text-zinc-400">{"{2}"}</span> for column 2.<br/>
                              Ex: <span className="text-zinc-400 italic">Page {"{1}"}: Verse {"{2}"}</span>
                            </div>
                          </div>
                          <input 
                            type="text"
                            value={state.indexFormat}
                            onChange={(e) => setState(prev => ({ ...prev, indexFormat: e.target.value }))}
                            placeholder="e.g. {1}:{2}"
                            className="w-full bg-white border border-zinc-100 rounded-xl px-4 py-2.5 text-[11px] font-bold shadow-sm focus:ring-1 focus:ring-black outline-none"
                          />
                        </div>

                        <div className="flex gap-2">
                          {(['content-bottom-right', 'footer'] as IndexPosition[]).map(pos => (
                            <button key={pos} onClick={() => setState(prev => ({ ...prev, indexPosition: pos }))} className={`flex-1 py-2.5 rounded-xl text-[9px] font-black uppercase tracking-tighter border transition-all ${state.indexPosition === pos ? 'bg-black text-white shadow-md' : 'bg-white text-zinc-300 border-zinc-100'}`}>{pos.replace(/-/g, ' ')}</button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </section>

              {/* 3. Template Designer Visual Section */}
              <section className="flex flex-col items-center border-t border-zinc-50 pt-12">
                <h3 className="text-[11px] font-black text-zinc-400 uppercase tracking-[0.4em] mb-8">3. Template Designer Canvas</h3>
                <div 
                  ref={previewRef}
                  className={`relative bg-white rounded-xl shadow-lg overflow-hidden border-[1px] border-black p-0 transition-all duration-700`}
                  style={{ 
                    width: '100%', 
                    maxWidth: '960px', 
                    aspectRatio: `${state.slideWidth}/${state.slideHeight}`
                  }}
                  onMouseDown={() => setSelectedDesignerBox(null)}
                >
                  <div className="absolute inset-0 bg-white">
                    {(Object.entries(state.mapping) as [keyof ColumnMapping, ColumnStyle][]).map(([key, style]) => {
                      if (!style.column) return null;
                      const isActive = activeSettingsField === key || hoveredField === key || selectedDesignerBox === key;
                      return (
                        <div 
                          key={key}
                          onMouseDown={(e) => startInteraction(e, key, 'move')}
                          style={{ 
                            left: `${style.x}%`, 
                            top: `${style.y}%`, 
                            width: `${style.w}%`, 
                            height: `${style.h}%`, 
                            zIndex: isActive ? 50 : 10 
                          }}
                          className={`absolute border-2 transition-all flex flex-col items-center justify-center cursor-move group ${isActive ? 'bg-zinc-100/60 border-black shadow-2xl z-50' : 'bg-zinc-50/10 border-zinc-200/50 hover:border-black'}`}
                        >
                          <div className="text-[9px] font-black text-white bg-black px-2 py-1 rounded shadow-sm uppercase tracking-tighter max-w-[90%] truncate">
                            {style.column}
                          </div>
                          
                          {isActive && (
                            <>
                              <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-zinc-400 opacity-20 pointer-events-none"><Move size={32} /></div>
                              <div 
                                onMouseDown={(e) => startInteraction(e, key, 'resize')}
                                className="absolute bottom-1 right-1 w-5 h-5 bg-black text-white rounded-md flex items-center justify-center cursor-se-resize shadow-md hover:scale-110 active:scale-90 transition-transform z-50 opacity-40 hover:opacity-100"
                              >
                                <Maximize size={10} />
                              </div>
                            </>
                          )}
                        </div>
                      );
                    })}
                    {state.includeIndex && (
                      <div className={`absolute border border-dashed border-zinc-200 text-[9px] font-black text-zinc-300 flex items-center justify-center uppercase tracking-widest ${state.indexPosition === 'content-bottom-right' ? 'right-6 bottom-12 w-20 h-5' : 'right-4 bottom-3 w-24 h-5'}`}>
                        INDEX
                      </div>
                    )}
                    <div className="absolute bottom-8 w-1/4 left-1/2 -translate-x-1/2 h-[1px] bg-zinc-50"></div>
                  </div>
                </div>
              </section>

              {/* 4. Filtering & Finish Controls */}
              <div className="mt-6 pt-10 border-t border-zinc-50">
                <h3 className="text-2xl font-black text-zinc-900 mb-8 flex items-center gap-3">
                  <List size={26} /> 4. Entry Filtering
                  <div className="group relative">
                    <HelpCircle size={16} className="text-zinc-300 cursor-help" />
                    <div className="absolute left-0 top-full mt-2 hidden group-hover:block bg-black text-white text-[10px] font-bold py-2 px-3 rounded-xl w-48 shadow-2xl z-50">
                      Choose which rows to generate.
                    </div>
                  </div>
                </h3>
                <div className="flex flex-wrap gap-4 mb-8">
                  {['all', 'range', 'specific'].map((mode) => (
                    <button 
                      key={mode} 
                      onClick={() => setState(prev => ({ ...prev, selectionMode: mode as RowSelectionMode }))} 
                      className={`px-10 py-5 rounded-[2rem] border font-black text-xs uppercase tracking-widest transition-all ${state.selectionMode === mode ? 'bg-black text-white border-black shadow-2xl scale-105' : 'bg-white text-zinc-300 border-zinc-100 hover:text-black hover:border-black'}`}
                    >
                      {mode === 'all' ? `All Entries (${state.data.length})` : mode}
                    </button>
                  ))}
                </div>

                {state.selectionMode !== 'all' && (
                  <div className="bg-zinc-50/20 p-8 rounded-[2.5rem] border border-zinc-50 max-w-xl mb-10 animate-in slide-in-from-top-4 duration-500">
                    {state.selectionMode === 'range' ? (
                      <div className="flex gap-8">
                        <div className="flex-1 space-y-3">
                          <label className="text-[10px] font-black text-zinc-300 uppercase tracking-widest block px-1">Start Row</label>
                          <input type="number" min="1" max={state.data.length} value={state.rangeStart} onChange={(e) => setState(prev => ({ ...prev, rangeStart: parseInt(e.target.value) || 1 }))} className="w-full bg-zinc-800 text-white border-none rounded-2xl px-5 py-4 text-sm font-bold shadow-lg focus:ring-2 focus:ring-white/20 outline-none" />
                        </div>
                        <div className="flex-1 space-y-3">
                          <label className="text-[10px] font-black text-zinc-300 uppercase tracking-widest block px-1">End Row</label>
                          <input type="number" min="1" max={state.data.length} value={state.rangeEnd} onChange={(e) => setState(prev => ({ ...prev, rangeEnd: parseInt(e.target.value) || state.data.length }))} className="w-full bg-zinc-800 text-white border-none rounded-2xl px-5 py-4 text-sm font-bold shadow-lg focus:ring-2 focus:ring-white/20 outline-none" />
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <label className="text-[10px] font-black text-zinc-300 uppercase tracking-widest block px-1">Specific row numbers</label>
                        <input type="text" placeholder="1, 4, 12..." value={state.specificRows} onChange={(e) => setState(prev => ({ ...prev, specificRows: e.target.value }))} className="w-full bg-zinc-800 text-white border-none rounded-2xl px-5 py-4 text-sm font-bold shadow-lg focus:ring-2 focus:ring-white/20 outline-none" />
                      </div>
                    )}
                  </div>
                )}

                <div className="flex justify-end pt-4">
                  <button 
                    onClick={() => setState(prev => ({ ...prev, step: AppStep.PREVIEW }))} 
                    disabled={!(Object.values(state.mapping) as ColumnStyle[]).some(m => m.column)} 
                    className="bg-black text-white font-black py-6 px-20 rounded-[2.5rem] flex items-center gap-3 hover:bg-zinc-800 disabled:bg-zinc-100 shadow-2xl transition-all transform hover:translate-x-2"
                  >
                    PROCEED TO REVIEW <ChevronRight size={22} />
                  </button>
                </div>
              </div>
            </div>
          )}

          {state.step === AppStep.PREVIEW && (
            <div className="animate-in zoom-in-95 py-10">
              <div className="flex items-center justify-between mb-16">
                <h2 className="text-3xl font-black text-zinc-900 tracking-tight">Final Deck Preview</h2>
                <button onClick={() => setState(prev => ({ ...prev, step: AppStep.MAPPING }))} className="text-zinc-400 hover:text-black font-black text-xs uppercase tracking-[0.2em] flex items-center gap-2 transition-colors"><ArrowLeft size={16} /> BACK TO DESIGNER</button>
              </div>

              <div
                ref={previewSlideRef}
                className="relative mx-auto rounded-[2.5rem] shadow-2xl overflow-hidden border border-zinc-200 mb-20 ring-4 ring-zinc-50"
                style={{
                  width: '100%',
                  maxWidth: '960px',
                  aspectRatio: `${state.slideWidth}/${state.slideHeight}`,
                  backgroundColor: themeBg
                }}
              >
                <div className="absolute inset-0">
                  {(Object.entries(state.mapping) as [keyof ColumnMapping, ColumnStyle][]).map(([key, style]) => {
                    if (!style.column || !filteredData[0]?.[style.column]) return null;
                    const val = String(filteredData[0][style.column]);
                    const isBox1 = key === 'box1';
                    return (
                      <div 
                        key={key} 
                        className={`absolute text-center flex items-center justify-center leading-relaxed ${isBox1 ? 'arabic-font' : ''}`}
                        style={{ 
                          left: `${style.x}%`, top: `${style.y}%`, width: `${style.w}%`, height: `${style.h}%`, 
                          fontSize: `${ptToPreviewPx(style.fontSize)}px`,
                          fontFamily: isBox1 ? 'Scheherazade New' : 'Arial',
                          color: themeText,
                          direction: isBox1 ? 'rtl' : 'ltr',
                          fontWeight: key === 'box2' ? 'bold' : 'normal',
                          fontStyle: key === 'box3' ? 'italic' : 'normal'
                        }}
                      >
                        {val}
                      </div>
                    );
                  })}
                  {state.includeIndex && (
                    <div
                      className="absolute font-black"
                      style={{
                        left: state.indexPosition === 'content-bottom-right' ? '75%' : '2%',
                        top: state.indexPosition === 'content-bottom-right' ? '90%' : '94%',
                        width: state.indexPosition === 'content-bottom-right' ? '20%' : '96%',
                        height: '5%',
                        fontFamily: 'Arial',
                        fontSize: `${ptToPreviewPx(state.indexPosition === 'content-bottom-right' ? 12 : 10)}px`,
                        color: themeText,
                        textAlign: 'right',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'flex-end',
                        paddingRight: ptToPreviewPx(4)
                      }}
                    >
                      {getPreviewIndex()}
                    </div>
                  )}
                </div>
              </div>

              <button onClick={handleGenerate} className="w-full bg-black text-white font-black py-7 rounded-[3rem] text-xl shadow-2xl hover:bg-zinc-800 transform hover:scale-[1.02] transition-all flex items-center justify-center gap-4">
                {state.isGenerating ? <div className="w-7 h-7 border-4 border-white/20 border-t-white rounded-full animate-spin"></div> : <Layers size={28} />}
                GENERATE {filteredData.length} SLIDES
              </button>
            </div>
          )}

          {state.step === AppStep.DOWNLOAD && (
            <div className="flex flex-col items-center text-center animate-in zoom-in-95 duration-700 py-24">
              <div className="w-32 h-32 bg-zinc-50 border border-zinc-100 rounded-[3.5rem] flex items-center justify-center mb-12 shadow-xl animate-bounce-subtle"><CheckCircle size={56} className="text-zinc-900" /></div>
              <h2 className="text-5xl font-black text-zinc-900 mb-5 tracking-tighter">Deck Exported</h2>
              <p className="text-zinc-400 mb-16 max-w-sm font-bold uppercase text-[11px] tracking-[0.4em] leading-loose">Automated grid-snapping applied to {filteredData.length} entries.</p>
              
              <div className="flex flex-col sm:flex-row gap-6 w-full max-w-lg">
                 <button onClick={handleGenerate} className="flex-1 bg-black text-white font-black py-6 rounded-[2.5rem] shadow-2xl hover:bg-zinc-800 flex items-center justify-center gap-3 tracking-[0.2em] uppercase text-xs transition-all hover:scale-105"><Download size={20} /> Download</button>
                 <button onClick={() => window.location.reload()} className="flex-1 bg-white border border-zinc-100 text-zinc-900 font-black py-6 rounded-[2.5rem] hover:bg-zinc-50 flex items-center justify-center gap-3 tracking-[0.2em] uppercase text-xs transition-all"><RotateCcw size={20} /> New Project</button>
              </div>
            </div>
          )}
        </div>
      </div>

      <footer className="mt-20 text-center text-zinc-300 text-[11px] font-black uppercase tracking-[0.5em] flex flex-col gap-4">
        <div className="flex justify-center items-center gap-6 opacity-40">
           <span className="flex items-center gap-2">XLSX COMPATIBLE</span>
           <span className="w-1.5 h-1.5 rounded-full bg-zinc-200"></span>
           <span className="flex items-center gap-2">CUSTOM DIMENSIONS</span>
           <span className="w-1.5 h-1.5 rounded-full bg-zinc-200"></span>
           <span className="flex items-center gap-2">SCHEHERAZADE FONT ENGINE</span>
        </div>
        <p className="opacity-30">© ARSLIDE DESIGN SYSTEMS</p>
      </footer>

      <style>{`
        @keyframes bounce-subtle {
          0%, 100% { transform: translateY(0); }
          50% { transform: translateY(-10px); }
        }
        .animate-bounce-subtle {
          animation: bounce-subtle 3s ease-in-out infinite;
        }
      `}</style>
    </div>
  );
};

export default App;
