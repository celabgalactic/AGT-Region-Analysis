/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, useRef, useCallback, ChangeEvent, ReactNode } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Search, 
  FileText, 
  Download, 
  Settings, 
  Database, 
  AlertCircle, 
  ChevronRight, 
  Table, 
  Columns,
  RefreshCw,
  Info,
  Volume2,
  VolumeX,
  Globe
} from 'lucide-react';
import Papa from 'papaparse';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { CIVILIZATIONS, GALAXIES } from './constants';
import { LanguageCode, LANGUAGES, TRANSLATIONS, translateColumnHeader } from './translations';
// @ts-ignore
import starSystemsIcon from './star-systems-icon.png';
// @ts-ignore
import agtAnthem from './AGT Anthem (Instrumental).mp3';

// Column configuration mapping
interface ColumnConfig {
  name: string;
  enabled: boolean;
  rawIndex: number;
}

type ReportType = 'simple' | 'detailed' | 'custom';

const colLetterToIdx = (letter: string): number => {
  let col = 0;
  letter.toUpperCase().trim().split('').forEach(char => {
    col = col * 26 + (char.charCodeAt(0) - 64);
  });
  return col - 1;
};

const parseExcelRange = (rangeStr: string): number[] => {
  const parts = rangeStr.toUpperCase().split(/\s+to\s+|\s*-\s*/);
  if (parts.length === 2) {
    const startIdx = colLetterToIdx(parts[0]);
    const endIdx = colLetterToIdx(parts[1]);
    const indices: number[] = [];
    for (let i = Math.min(startIdx, endIdx); i <= Math.max(startIdx, endIdx); i++) {
      indices.push(i);
    }
    return indices;
  }
  return [colLetterToIdx(rangeStr)];
};

const CUSTOM_COLUMN_DEFS = [
  { id: 'A', label: 'Galaxy', letters: 'A' },
  { id: 'B', label: 'Region', letters: 'B' },
  { id: 'C', label: 'System Name', letters: 'C' },
  { id: 'G', label: 'Original name', letters: 'G' },
  { id: 'H', label: 'Coordinates', letters: 'H' },
  { id: 'I', label: 'Glyphs', letters: 'I' },
  { id: 'K', label: 'Survey', letters: 'K' },
  { id: 'L', label: 'Discoverer', letters: 'L' },
  { id: 'P', label: 'Survey Date', letters: 'P' },
  { id: 'O', label: 'Discovery Date', letters: 'O' },
  { id: 'N', label: 'Giant?', letters: 'N' },
  { id: 'Q', label: 'BH/Atlas', letters: 'Q' },
  { id: 'R', label: 'Dissonant', letters: 'R' },
  { id: 'S', label: 'Civ', letters: 'S' },
  { id: 'V', label: 'Platform', letters: 'V' },
  { id: 'W', label: 'Mode', letters: 'W' },
  { id: 'X', label: 'Stars', letters: 'X' },
  { id: 'Y', label: 'Category', letters: 'Y' },
  { id: 'Z', label: 'Color', letters: 'Z' },
  { id: 'AA', label: 'Planets', letters: 'AA' },
  { id: 'AB', label: 'Moons', letters: 'AB' },
  { id: 'AC', label: 'Faction', letters: 'AC' },
  { id: 'AE', label: 'Distance', letters: 'AE' },
  { id: 'AF', label: 'Water', letters: 'AF' },
  { id: 'AG', label: 'Economy', letters: 'AG' },
  { id: 'AH', label: 'Wealth', letters: 'AH' },
  { id: 'AI', label: 'Ebuy', letters: 'AI' },
  { id: 'AJ', label: 'ESell', letters: 'AJ' },
  { id: 'AK', label: 'Conflict', letters: 'AK' },
  { id: 'AL', label: 'Release', letters: 'AL' },
  { id: 'CQ', label: 'Rel#', letters: 'CQ' },
  { id: 'AN_AR', label: 'Trade', letters: 'AN to AR' },
  { id: 'AS_BG', label: 'Updates', letters: 'AS to BG' },
  { id: 'BM_BU', label: 'Notes', letters: 'BM to BU' },
  { id: 'BV', label: 'Phantom', letters: 'BV' },
  { id: 'BW', label: 'CTR Access', letters: 'BW' },
  { id: 'BY', label: 'Wiki Link', letters: 'BY' },
  { id: 'BZ_CD', label: 'Other Links', letters: 'BZ to CD' },
  { id: 'CE_CJ', label: 'Legacy Info', letters: 'CE to CJ' },
  { id: 'CL', label: 'Age', letters: 'CL' },
  { id: 'CM', label: 'Research', letters: 'CM' },
  { id: 'CN', label: 'Misc', letters: 'CN' },
  { id: 'CY', label: 'Wealth Lvl', letters: 'CY' },
  { id: 'CZ', label: 'Conflict Lvl', letters: 'CZ' }
];

const CIV_ACRONYMS: Record<string, string> = {
  'Alliance of Galactic Travellers': 'AGT',
  'Intergalactic Travellers Foundation': 'IGTF',
  'Calypso Travellers Foundation': 'CTF',
  'Hyades Travellers Foundation': 'HTF',
  'Budullanger Travellers Foundation': 'BTF',
  'Budullangr Travellers Foundation': 'BTF',
  'Isdoraijung Travellers Foundation': 'ITF',
  'Kikolgallr Travellers Foundation': 'KTF',
  'Eissentam Travellers Foundation': 'ETF',
  'Ickjamatew Travellers Foundation': 'IJTF',
  'Rycempler Travellers Foundation': 'RTF',
  'Zavainlani Travellers Foundation': 'ZTF',
  'Animal Cracker Projects': 'ACP',
  'United Star Navy': 'USN',
  'CELAB Galactic Industries': 'CGI',
  'IVc Project': 'IVc',
  'AAAM Expeditionary': 'AAAM',
  'Riven Minerals and Exploration': 'RME',
  'Gravemind Expeditionary Force': 'GMEF'
};

const getDisplayValue = (val: any, colIdx?: number) => {
  const strVal = String(val || '').trim();
  if (colIdx === 18 && CIV_ACRONYMS[strVal]) {
    return CIV_ACRONYMS[strVal];
  }
  return strVal;
};

interface AutocompleteProps {
  value: string;
  onChange: (val: string) => void;
  suggestions: string[];
  placeholder: string;
  id: string;
  icon: ReactNode;
}

function AutocompleteInput({ value, onChange, suggestions, placeholder, id, icon }: AutocompleteProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [filterText, setFilterText] = useState(value);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setFilterText(value);
  }, [value]);

  const filteredSuggestions = useMemo(() => {
    const text = filterText.trim().toLowerCase();
    const pool = Array.from(new Set(suggestions)).filter(s => s.toLowerCase() !== 'all');
    if (!text) return pool.slice(0, 10);
    return pool
      .filter(s => s.toLowerCase().includes(text) && s.toLowerCase() !== text)
      .slice(0, 10);
  }, [filterText, suggestions]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div ref={containerRef} className="relative w-full group">
      <div className="absolute inset-y-0 left-0 pl-6 flex items-center pointer-events-none text-[#FFB451] transition-colors">
        {icon}
      </div>
      <input
        id={id}
        type="text"
        value={filterText}
        placeholder={placeholder}
        onFocus={() => setIsOpen(true)}
        onChange={(e) => {
          const val = e.target.value;
          setFilterText(val);
          onChange(val);
        }}
        className="block w-full pl-14 pr-12 py-5 bg-[#1d1d1d] border-2 border-[#FF0500] rounded-full text-lg font-mono tracking-wider focus:outline-none focus:ring-2 focus:ring-[#FF0500] focus:border-[#FF0500] transition-all input-glow text-[#FFB451] shadow-[0_0_30px_rgba(255,5,0,0.05)] placeholder-[#FFB451]/30"
      />
      
      <AnimatePresence>
        {isOpen && filteredSuggestions.length > 0 && (
          <motion.div
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="absolute z-[80] left-0 right-0 mt-2 bg-[#161616] border border-[#FF0500] rounded-2xl overflow-hidden shadow-[0_10px_30px_rgba(255,5,0,0.4)] max-h-60 overflow-y-auto settings-scrollbar"
          >
            {filteredSuggestions.map((suggestion) => (
              <button
                key={suggestion}
                type="button"
                onMouseDown={() => {
                  onChange(suggestion);
                  setFilterText(suggestion);
                  setIsOpen(false);
                }}
                className="w-full text-left px-6 py-3 text-sm font-mono text-white hover:bg-[#E25530] hover:text-white transition-colors border-b border-[#FF0500]/10 last:border-0"
              >
                {suggestion}
              </button>
            ))}
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

const getSlicePath = (
  cx: number,
  cy: number,
  r: number,
  startAngle: number,
  endAngle: number
) => {
  const startRad = (startAngle - 90) * Math.PI / 180;
  const endRad = (endAngle - 90) * Math.PI / 180;
  
  const x1 = cx + r * Math.cos(startRad);
  const y1 = cy + r * Math.sin(startRad);
  const x2 = cx + r * Math.cos(endRad);
  const y2 = cy + r * Math.sin(endRad);
  
  const largeArcFlag = endAngle - startAngle <= 180 ? 0 : 1;
  const pathData = [
    `M ${cx} ${cy}`,
    `L ${x1} ${y1}`,
    `A ${r} ${r} 0 ${largeArcFlag} 1 ${x2} ${y2}`,
    `Z`
  ].join(' ');
  
  return pathData;
};

interface PieDataPoint {
  label: string;
  value: number;
  color?: string;
}

function SvgPieChart({ title, data, valueLabel = 'count', useSpecialColors = false, hideLegend = false, radius = 65, height = 220, labelFontSize = 8, forceOppositeSides = false, hideSliceLabels = false }: {
  title: string;
  data: PieDataPoint[];
  valueLabel?: string;
  useSpecialColors?: boolean;
  hideLegend?: boolean;
  radius?: number;
  height?: number;
  labelFontSize?: number;
  forceOppositeSides?: boolean;
  hideSliceLabels?: boolean;
}) {
  const total = data.reduce((sum, item) => sum + item.value, 0);
  if (total === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-64 border border-[#FF0500]/10 rounded-xl bg-black/40">
        <span className="text-xs text-white/40 uppercase tracking-widest">{title} - No Data</span>
      </div>
    );
  }

  const colors = [
    '#FFB451', // AGT Orange
    '#EF4444', // Red
    '#3B82F6', // Blue
    '#10B981', // Emerald Green
    '#8B5CF6', // Purple
    '#F59E0B', // Amber
    '#06B6D4', // Cyan
    '#EC4899', // Pink
    '#14B8A6', // Teal
    '#6366F1'  // Indigo
  ];

  const slices: { label: string; value: number; percent: number; startAngle: number; endAngle: number; color: string }[] = [];
  let currentAngle = 0;

  data.forEach((item, index) => {
    const percent = item.value / total;
    const angle = percent * 360;
    
    let color = item.color || colors[index % colors.length];
    
    if (useSpecialColors) {
      if (item.label === 'Alliance of Galactic Travellers') {
        color = '#FFB451';
      }
    }

    slices.push({
      label: item.label,
      value: item.value,
      percent: percent,
      startAngle: currentAngle,
      endAngle: currentAngle + angle,
      color: color
    });
    
    currentAngle += angle;
  });

  const cx = 225;
  const cy = 130;
  const r = radius;

  return (
    <div className="flex flex-col items-center my-6 py-8 px-6 bg-black/30 border border-[#FF0500]/10 rounded-2xl relative shadow-[0_0_15px_rgba(255,5,0,0.02)] h-full justify-between">
      <h4 style={{ fontSize: `${labelFontSize + 1}px` }} className="font-bold uppercase tracking-widest text-[#FFB451]/90 mb-2 self-start border-l-2 border-[#FF0500] pl-2">{title}</h4>
      <div className="w-full relative shrink-0" style={{ height: `${height}px` }}>
        <svg viewBox="0 0 450 240" className="w-full h-full overflow-visible">
          <circle cx={cx} cy={cy} r={r + 5} fill="none" stroke="#FF0500" strokeWidth="1" strokeOpacity="0.05" />

          {slices.map((slice, i) => {
            if (slice.percent >= 0.999) {
              return (
                <g key={i}>
                  <circle cx={cx} cy={cy} r={r} fill={slice.color} />
                  {!hideSliceLabels && (
                    <>
                      <line x1={cx} y1={cy - r} x2={cx} y2={cy - r - 20} stroke={slice.color} strokeWidth="1.5" />
                      <line x1={cx} y1={cy - r - 20} x2={cx + 15} y2={cy - r - 20} stroke={slice.color} strokeWidth="1.5" />
                      <text 
                        x={cx + 20} 
                        y={cy - r - 20} 
                        alignmentBaseline="middle" 
                        textAnchor="start" 
                        fill="#FFB451" 
                        style={{ fontSize: `${labelFontSize}px` }}
                        className="font-mono uppercase tracking-wider font-semibold"
                      >
                        {slice.label} [{Math.round(slice.percent * 100)}%]
                      </text>
                    </>
                  )}
                </g>
              );
            }

            const pathData = getSlicePath(cx, cy, r, slice.startAngle, slice.endAngle);
            const midAngle = slice.startAngle + (slice.endAngle - slice.startAngle) / 2;
            const midRad = (midAngle - 90) * Math.PI / 180;
            
            const px1 = cx + (r * 0.8) * Math.cos(midRad);
            const py1 = cy + (r * 0.8) * Math.sin(midRad);
            const px2 = cx + (r * 1.2) * Math.cos(midRad);
            const py2 = cy + (r * 1.2) * Math.sin(midRad);
            
            let actualPx2 = px2;
            let actualPy2 = py2;
            if (forceOppositeSides) {
              const labelLower = slice.label.toLowerCase();
              if (labelLower.includes('no')) {
                actualPx2 = cx - r - 65;
                actualPy2 = cy;
              } else if (labelLower.includes('yes')) {
                actualPx2 = cx + r + 65;
                actualPy2 = cy - 35;
              } else {
                actualPx2 = cx + r + 65;
                actualPy2 = cy + 35;
              }
            }

            const isRight = actualPx2 > cx;
            const px3 = actualPx2 + (isRight ? 15 : -15);
            const showPointer = forceOppositeSides ? true : (slice.percent >= 0.015);

            return (
              <g key={i} className="hover:opacity-90 transition-opacity cursor-pointer group">
                <path d={pathData} fill={slice.color} />
                
                {showPointer && !hideSliceLabels && (
                  <g>
                    <path d={`M ${px1} ${py1} L ${actualPx2} ${actualPy2} L ${px3} ${actualPy2}`} fill="none" stroke={slice.color} strokeWidth="1.0" strokeOpacity="0.8" />
                    <circle cx={actualPx2} cy={actualPy2} r="1.5" fill={slice.color} />
                    <text 
                      x={px3 + (isRight ? 3 : -3)} 
                      y={actualPy2} 
                      alignmentBaseline="middle" 
                      textAnchor={isRight ? "start" : "end"} 
                      fill="#FFB451" 
                      style={{ fontSize: `${labelFontSize}px` }}
                      className="font-mono uppercase tracking-widest font-semibold drop-shadow-[0_1px_4px_rgba(0,0,0,0.6)]"
                    >
                      {slice.label} [{Math.round(slice.percent * 100)}%]
                    </text>
                  </g>
                )}
                <title>{slice.label}: {slice.value} ({Math.round(slice.percent * 100)}%)</title>
              </g>
            );
          })}
        </svg>
      </div>

      {!hideLegend && (
        <div className="flex flex-wrap gap-x-3 gap-y-1 justify-center mt-2 max-h-14 overflow-y-auto w-full border-t border-[#FF0500]/5 pt-2 settings-scrollbar shrink-0">
          {slices.map((slice, i) => (
            <div key={i} className="flex items-center gap-1 text-[8px] font-mono uppercase tracking-wider text-white/60">
              <span className="w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
              <span>{slice.label} ({slice.value})</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const parseDateSafe = (dateStr: any): Date => {
  if (!dateStr) return new Date(0);
  const str = String(dateStr).trim();
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) return new Date(parsed);
  
  const parts = str.split(/[-./]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const y = parseInt(parts[2], 10);
    if (y < 100) {
      const fullY = y + (y > 30 ? 1900 : 2000);
      return new Date(fullY, m - 1, d);
    }
    return new Date(y, m - 1, d);
  }
  return new Date(0);
};

const parseTradeValue = (val: any): number => {
  if (val === undefined || val === null) return 0;
  const str = String(val).trim();
  const cleaned = str.replace(/[^\d.-]/g, '');
  if (!cleaned) return 0;
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
};

const getCellValue = (row: any, colLetter: string): string => {
  if (!row || !row._rawRow) return '';
  const idx = colLetterToIdx(colLetter);
  let val = String(row._rawRow[idx] || '').trim();
  if (colLetter.toUpperCase().trim() === 'C' && val) {
    const bvIdx = colLetterToIdx('BV');
    const bvVal = String(row._rawRow[bvIdx] || '').trim().toLowerCase();
    if (bvVal === 'phantom') {
      val = val + ' [Phantom]';
    } else if (bvVal === 'shadow') {
      val = val + ' [Shadow]';
    } else if (bvVal === 'reverse shadow') {
      val = val + ' [Reverse]';
    }
  }
  return val;
};

export default function App() {
  const [reportType, setReportType] = useState<ReportType>('simple');
  const [sheetUrl, setSheetUrl] = useState<string>(() => {
    const saved = localStorage.getItem('sheet_reporter_url');
    const oldDefault = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSWiJE26JMTHgjGeZfpfTrwT1HL2ZnXIqiOVkNs-V8wtDkGE7ey0Q9hnAM-bpMhy475q45qHa09o2vC/pub?gid=0&single=true&output=csv';
    const previousDefault = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0jFq80ut0o5jtApdhRG8sR2CIufVn0FNcugR_7fdCIfrDRfgB9s-SvEhBAePrQCibr1RcxFVoXj7o/pub?gid=354119689&single=true&output=tsv';
    const newDefault = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0jFq80ut0o5jtApdhRG8sR2CIufVn0FNcugR_7fdCIfrDRfgB9s-SvEhBAePrQCibr1RcxFVoXj7o/pub?gid=0&single=true&output=tsv';
    
    if (!saved || saved === oldDefault || saved === previousDefault) return newDefault;
    return saved;
  });
  const [showSettings, setShowSettings] = useState(false);
  const [audioEnabled, setAudioEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem('agt_audio_enabled');
    return saved === null ? true : saved === 'true';
  });
  
  const audioRef = useRef<HTMLAudioElement>(null);

  // Initial fetch
  useEffect(() => {
    if (sheetUrl) {
      fetchData();
    }
  }, []);

  // Background Audio Management
  useEffect(() => {
    const handleFirstInteraction = () => {
      if (audioEnabled && audioRef.current) {
        audioRef.current.volume = 0.4;
        audioRef.current.play().catch(() => {});
      }
      window.removeEventListener('mousedown', handleFirstInteraction);
      window.removeEventListener('keydown', handleFirstInteraction);
      window.removeEventListener('touchstart', handleFirstInteraction);
    };

    window.addEventListener('mousedown', handleFirstInteraction);
    window.addEventListener('keydown', handleFirstInteraction);
    window.addEventListener('touchstart', handleFirstInteraction);

    return () => {
      window.removeEventListener('mousedown', handleFirstInteraction);
      window.removeEventListener('keydown', handleFirstInteraction);
      window.removeEventListener('touchstart', handleFirstInteraction);
    };
  }, [audioEnabled]);

  useEffect(() => {
    if (audioRef.current) {
      audioRef.current.volume = 0.4;
      if (audioEnabled) {
        audioRef.current.play().catch(() => {});
      } else {
        audioRef.current.pause();
      }
    }
    localStorage.setItem('agt_audio_enabled', String(audioEnabled));
  }, [audioEnabled]);

  const handleManualPlay = () => {
    if (audioEnabled && audioRef.current && audioRef.current.paused) {
      audioRef.current.play().catch(() => {});
    }
  };

  const [logoUrl, setLogoUrl] = useState('/AGTicon.png');
  const [logoTriedCount, setLogoTriedCount] = useState(0);
  const [pageSize, setPageSize] = useState<number>(() => {
    const saved = localStorage.getItem('agt_page_size');
    return saved ? parseInt(saved, 10) : 15;
  });
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [fontScale, setFontScale] = useState<string>(() => {
    return localStorage.getItem('agt_font_scale') || '1x';
  });
  const [sortColumn, setSortColumn] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc' | null>(null);

  useEffect(() => {
    localStorage.setItem('agt_page_size', String(pageSize));
  }, [pageSize]);

  useEffect(() => {
    localStorage.setItem('agt_font_scale', fontScale);
  }, [fontScale]);

  const [language, setLanguage] = useState<LanguageCode>(() => {
    return (localStorage.getItem('agt_app_language') as LanguageCode) || 'en';
  });

  useEffect(() => {
    localStorage.setItem('agt_app_language', language);
  }, [language]);

  const t = useCallback((key: string): string => {
    if (language === 'en') return key;
    const langSet = TRANSLATIONS[language];
    if (langSet && langSet[key]) {
      return langSet[key];
    }
    return key;
  }, [language]);

  const [searchKey, setSearchKey] = useState('All');
  const [selectedGalaxy, setSelectedGalaxy] = useState('All');
  const [selectedRegion, setSelectedRegion] = useState('');
  const [popupError, setPopupError] = useState<string | null>(null);
  const [discovererName, setDiscovererName] = useState('');
  const [surveyorName, setSurveyorName] = useState('');
  const [generationSeconds, setGenerationSeconds] = useState(0);
  const [pdfErrorMsg, setPdfErrorMsg] = useState<string | null>(null);
  const [searchLoading, setSearchLoading] = useState(false);
  const [reportGeneratingLoading, setReportGeneratingLoading] = useState(false);

  useEffect(() => {
    let interval: any = null;
    if (reportGeneratingLoading) {
      setGenerationSeconds(0);
      interval = setInterval(() => {
        setGenerationSeconds(prev => prev + 1);
      }, 1000);
    } else {
      setGenerationSeconds(0);
    }
    return () => {
      if (interval) clearInterval(interval);
    };
  }, [reportGeneratingLoading]);

  const [enabledCustomColumns, setEnabledCustomColumns] = useState<Record<string, boolean>>(() => {
    const saved = localStorage.getItem('agt_custom_columns_toggles');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        // ignore
      }
    }
    const defaults: Record<string, boolean> = {};
    CUSTOM_COLUMN_DEFS.forEach(def => {
      defaults[def.id] = true;
    });
    return defaults;
  });

  useEffect(() => {
    localStorage.setItem('agt_custom_columns_toggles', JSON.stringify(enabledCustomColumns));
  }, [enabledCustomColumns]);

  const [data, setData] = useState<any[]>([]);

  const civSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[18] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set([...CIVILIZATIONS, ...rawList]));
    return unique.sort();
  }, [data]);

  const galaxySuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[0] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set([...GALAXIES, ...rawList]));
    return unique.sort();
  }, [data]);

  const regionSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[1] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set(rawList));
    return unique.filter(r => String(r).toLowerCase() !== 'all').sort();
  }, [data]);

  const discovererSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[11] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set(rawList));
    return unique.sort();
  }, [data]);

  const surveyorSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[10] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set(rawList));
    return unique.sort();
  }, [data]);

  const [columns, setColumns] = useState<ColumnConfig[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [matchedRecords, setMatchedRecords] = useState<any[]>([]);

  // Save sheet URL to localStorage
  useEffect(() => {
    if (sheetUrl) {
      localStorage.setItem('sheet_reporter_url', sheetUrl);
    }
  }, [sheetUrl]);

  useEffect(() => {
    setCurrentPage(1);
  }, [matchedRecords, pageSize]);

  const fetchData = async () => {
    if (!sheetUrl) {
      setError('Please provide a Google Sheet CSV URL in settings.');
      setShowSettings(true);
      return;
    }

    setLoading(true);
    setSearchLoading(true);
    const startTime = Date.now();
    setError(null);
    setMatchedRecords([]);

    try {
      // Handle the case where the user might paste a regular sheet URL instead of a pub link
      let fetchUrl = sheetUrl;
      if (sheetUrl.includes('docs.google.com/spreadsheets/') && !sheetUrl.includes('pub?')) {
        // Try to convert regular URL to CSV export if possible, 
        // though "Publish to Web" is the official way.
        if (sheetUrl.includes('/edit')) {
          fetchUrl = sheetUrl.replace(/\/edit.*$/, '/export?format=csv');
        }
      }

      const response = await fetch(fetchUrl);
      if (!response.ok) throw new Error('Failed to fetch sheet data. Is it published to the web?');
      
      const csvText = await response.text();
      
      Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true,
        delimiter: fetchUrl.includes('output=tsv') ? '\t' : undefined,
        complete: (results) => {
          const rawRows = results.data as string[][];
          if (rawRows.length < 2) {
            setError('The source sheet data is insufficient (need at least 2 rows).');
            const elapsed = Date.now() - startTime;
            setTimeout(() => {
              setLoading(false);
              setSearchLoading(false);
            }, Math.max(0, 1500 - elapsed));
            return;
          }

          const headers = rawRows[1]; // Row 2 is headers
          
          const simpleIndices = [0, 1, 2, 6, 7, 10, 11, 14, 15, 18, 76];
          const detailedIndices = [0, 1, 2, 6, 7, 8, 10, 11, 13, 14, 15, 17, 18, 22, 23, 24, 25, 26, 27, 28, 31, 32, 33, 34, 35, 36, 37, 76];
          
          let targetIndexes: number[];
          if (reportType === 'simple') {
            targetIndexes = simpleIndices;
          } else if (reportType === 'detailed') {
            targetIndexes = detailedIndices;
          } else {
            const indicesSet = new Set<number>();
            CUSTOM_COLUMN_DEFS.forEach(def => {
              if (enabledCustomColumns[def.id]) {
                parseExcelRange(def.letters).forEach(idx => indicesSet.add(idx));
              }
            });
            targetIndexes = Array.from(indicesSet).sort((a, b) => a - b);
            if (targetIndexes.length === 0) {
              targetIndexes = simpleIndices;
            }
          }
          
          const filteredColumns = targetIndexes.map(idx => ({
            name: headers[idx] || `Col ${String.fromCharCode(65 + (idx % 26))}${idx >= 26 ? String.fromCharCode(65 + Math.floor(idx / 26) - 1) : ''}`,
            enabled: true,
            rawIndex: idx
          }));
          
          setColumns(filteredColumns);
          
          const processedData = rawRows.slice(3) // Skip Rows 1, 2, 3 (index 0, 1, 2)
            .filter(row => {
              const colA = String(row[0] || '').trim();
              const colB = String(row[1] || '').trim();
              
              // Skip if:
              // - Col A is empty
              // - Col A has SKIPROW
              // - Col A has #N/A
              // - Col B is blank
              if (colA === '' || colA.includes('SKIPROW') || colA.includes('#N/A')) return false;
              if (colB === '') return false;
              
              // Exclude all records with value 99 in column AM (index 38)
              const colAM = String(row[38] || '').trim();
              if (colAM === '99') return false;
              
              return true;
            })
            .map(row => {
              const rowObj: any = { _rawRow: row };
              targetIndexes.forEach((colIdx, listIdx) => {
                const headerName = filteredColumns[listIdx].name;
                rowObj[headerName] = row[colIdx] || '';
              });
              return rowObj;
            });
          
          setData(processedData);
          
          const regionList = processedData.map(row => String(row._rawRow ? row._rawRow[1] : '').trim()).filter(Boolean);
          const uniqueRegions = Array.from(new Set(regionList)).sort();
          let regionToScan = selectedRegion.trim();
          
          if (regionToScan !== '' && regionToScan.toLowerCase() !== 'all') {
            findRecord(processedData, filteredColumns, regionToScan);
          } else {
            setMatchedRecords([]);
          }
          
          const elapsed = Date.now() - startTime;
          setTimeout(() => {
            setLoading(false);
            setSearchLoading(false);
          }, Math.max(0, 1500 - elapsed));
        },
        error: (err: any) => {
          setError(`Parsing error: ${err.message}`);
          const elapsed = Date.now() - startTime;
          setTimeout(() => {
            setLoading(false);
            setSearchLoading(false);
          }, Math.max(0, 1500 - elapsed));
        }
      });
    } catch (err: any) {
      setError(err.message || 'Operation failed');
      const elapsed = Date.now() - startTime;
      setTimeout(() => {
        setLoading(false);
        setSearchLoading(false);
      }, Math.max(0, 1500 - elapsed));
    }
  };

  const handleSearch = async () => {
    const trimmed = selectedRegion.trim();
    if (!trimmed || trimmed.toLowerCase() === 'all') {
      setPopupError("User must enter a valid region name");
      return;
    }
    setSearchLoading(true);
    await fetchData();
  };

  const findRecord = (
    sourceData: any[], 
    sourceCols: ColumnConfig[], 
    regionTerm?: string
  ) => {
    const currentRegionTerm = (regionTerm ?? selectedRegion).trim().toLowerCase();
    
    const getRawVal = (r: any, rawIdx: number) => {
      if (r._rawRow) return String(r._rawRow[rawIdx] || '').trim();
      const colMatch = sourceCols.find(c => c.rawIndex === rawIdx);
      return colMatch ? String(r[colMatch.name] || '').trim() : '';
    };

    const matches = sourceData.filter(row => {
      const rVal = getRawVal(row, 1);  // Region is Column B (index 1)

      // Region matches: must match exact currentRegionTerm (no ALL support)
      if (!currentRegionTerm) {
        return false;
      } else {
        return rVal.toLowerCase() === currentRegionTerm;
      }
    });

    // Sort by Column A then Column B
    const sortedMatches = [...matches].sort((a, b) => {
      const galA = getRawVal(a, 0).toLowerCase();
      const galB = getRawVal(b, 0).toLowerCase();
      if (galA !== galB) return galA.localeCompare(galB);
      
      const regionA = getRawVal(a, 1).toLowerCase();
      const regionB = getRawVal(b, 1).toLowerCase();
      return regionA.localeCompare(regionB);
    });

    if (sortedMatches.length > 0) {
      setMatchedRecords(sortedMatches);
      setError(null);
    } else {
      setMatchedRecords([]);
      setError(`No records found for the selected region.`);
    }
  };

  const formatTimestamp = () => {
    const now = new Date();
    const pad = (num: number) => String(num).padStart(2, '0');
    return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}_${pad(now.getMinutes())}_${pad(now.getSeconds())}`;
  };

  const generatePieChartDataUrl = (data: { label: string; value: number; color?: string }[]) => {
    const canvas = document.createElement('canvas');
    canvas.width = 300;
    canvas.height = 300;
    const ctx = canvas.getContext('2d');
    if (!ctx) return '';

    ctx.clearRect(0, 0, 300, 300);

    const total = data.reduce((sum, item) => sum + item.value, 0);
    if (total === 0) {
      ctx.strokeStyle = '#cccccc';
      ctx.lineWidth = 4;
      ctx.beginPath();
      ctx.arc(150, 150, 120, 0, 2 * Math.PI);
      ctx.stroke();
      return canvas.toDataURL('image/png');
    }

    let startAngle = 0;
    data.forEach(item => {
      const sliceAngle = (item.value / total) * 2 * Math.PI;
      if (sliceAngle > 1e-5) {
        ctx.fillStyle = item.color || '#A855F7';
        ctx.beginPath();
        ctx.moveTo(150, 150);
        ctx.arc(150, 150, 120, startAngle, startAngle + sliceAngle);
        ctx.closePath();
        ctx.fill();
        
        ctx.strokeStyle = '#ffffff';
        ctx.lineWidth = 1;
        ctx.stroke();
        
        startAngle += sliceAngle;
      }
    });

    return canvas.toDataURL('image/png');
  };

  const downloadRegionAnalysisReportPdf = () => {
    if (!analysisStats) return;
    setReportGeneratingLoading(true);

    setTimeout(() => {
      try {
        const doc = new jsPDF('p', 'mm', 'a4');
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const now = new Date();
        const browserDateStr = `${String(now.getDate()).padStart(2, '0')}-${months[now.getMonth()]}-${now.getFullYear()}`;
        
        const getFormattedSystemDateTime = (date: Date) => {
          const day = String(date.getDate()).padStart(2, '0');
          const month = months[date.getMonth()];
          const year = date.getFullYear();
          const pad = (num: number) => String(num).padStart(2, '0');
          const timeStr = `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
          return `${day}-${month}-${year} ${timeStr}`;
        };

        const drawSectionHeading = (docInstance: typeof doc, title: string, y: number) => {
          docInstance.setFont("Helvetica", "bold");
          docInstance.setFontSize(13);
          docInstance.setTextColor(226, 85, 48); // #E25530
          docInstance.text(title, 15, y);
          
          docInstance.setDrawColor(226, 85, 48);
          docInstance.setLineWidth(0.5);
          docInstance.line(15, y + 2, 195, y + 2);
          return y + 8;
        };

        const drawSubsectionHeading = (docInstance: typeof doc, title: string, y: number) => {
          docInstance.setFont("Helvetica", "bold");
          docInstance.setFontSize(10);
          docInstance.setTextColor(0, 0, 0);
          docInstance.text(title, 15, y);
          return y + 4;
        };

        // Helper to add a page and draw the header/footer
        const createNewPage = (reportPageNum: number) => {
          doc.addPage();
          
          // Header Logo (AGTiconLogo) left-aligned
          try {
            const imgElements = document.getElementsByTagName('img');
            let logoImg: HTMLImageElement | null = null;
            for (let i = 0; i < imgElements.length; i++) {
              if (imgElements[i].alt === 'AGT Logo') {
                logoImg = imgElements[i];
                break;
              }
            }
            if (logoImg && logoImg.complete) {
              doc.addImage(logoImg, 'PNG', 15, 8, 12, 12);
            } else {
              const tempImg = new Image();
              tempImg.src = logoUrl;
              if (tempImg.complete) {
                doc.addImage(tempImg, 'PNG', 15, 8, 12, 12);
              }
            }
          } catch (err) {
            console.warn('Header logo failed:', err);
          }
          
          doc.setFont("Helvetica", "bold");
          doc.setFontSize(10);
          doc.setTextColor(0, 0, 0);
          doc.text(`AGT Region Analysis: ${selectedRegion}`, 32, 15);
          
          // Page Number
          doc.text(`${reportPageNum}`, 195, 15, { align: 'right' });
          
          // Underline header
          doc.setDrawColor(220, 220, 220);
          doc.setLineWidth(0.3);
          doc.line(15, 23, 195, 23);
          
          // Footer starting at Page 2 (reportPageNum >= 2)
          if (reportPageNum >= 2) {
            doc.setDrawColor(220, 220, 220);
            doc.setLineWidth(0.3);
            doc.line(15, 280, 195, 280);
            
            doc.setFont("Helvetica", "normal");
            doc.setFontSize(8);
            doc.setTextColor(100, 100, 100);
            const dateStr = getFormattedSystemDateTime(new Date());
            doc.text(`Date of Report: ${dateStr}`, 15, 287);
          }
        };

        // --- 1. COVER PAGE (PDF Page 1) ---
        // 'AGTIcon.png' horizontally centered and 20% from top margin (y = 59.4)
        try {
          const imgElements = document.getElementsByTagName('img');
          let logoImg: HTMLImageElement | null = null;
          for (let i = 0; i < imgElements.length; i++) {
            if (imgElements[i].alt === 'AGT Logo') {
              logoImg = imgElements[i];
              break;
            }
          }
          if (logoImg && logoImg.complete) {
            doc.addImage(logoImg, 'PNG', 90, 59.4, 30, 30);
          } else {
            const tempImg = new Image();
            tempImg.src = logoUrl;
            if (tempImg.complete) {
              doc.addImage(tempImg, 'PNG', 90, 59.4, 30, 30);
            }
          }
        } catch (err) {
          console.warn('Cover logo failed:', err);
        }

        // Title text "AGT Region Analysis" centered horizontally in FF0500 hex.
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(26);
        doc.setTextColor(255, 5, 0); // RGB hex FF0500
        doc.text("AGT Region Analysis", 105, 105, { align: 'center' });

        // Below the title, text in black
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(12);
        doc.setTextColor(0, 0, 0);
        doc.text(`Region: ${selectedRegion}`, 105, 122, { align: 'center' });
        
        doc.setFont("Helvetica", "normal");
        doc.text(`Date of Report: ${browserDateStr}`, 105, 129, { align: 'center' });

        // --- REPORT PAGE 1 (PDF Page 2) ---
        createNewPage(1);
        let currentY = 30;

        // Section 1 Heading
        currentY = drawSectionHeading(doc, "SECTION 1: REGION CHARACTERISTICS", currentY);
        
        // Basic Facts
        currentY = drawSubsectionHeading(doc, "Basic Facts", currentY);
        
        const basicFactsBody = [
          ["Region Name:", selectedRegion],
          ["Region Age:", analysisStats.regionAge],
          ["Region Claim Status:", analysisStats.regionClaimStatus],
          ["Estimated Normal Systems:", analysisStats.estimatedNormalSystems],
          ["Galactic Coordinate Base Address:", analysisStats.coordinateBaseAddress],
          ["Galaxy:", analysisStats.galaxyName],
          ["Special: Black Hole System:", analysisStats.blackHoleSystem || ' - '],
          ["Special: Atlas Station System:", analysisStats.atlasSystem || ' - '],
          ["Special: Shadow Star System:", analysisStats.shadowSystem || ' - '],
          ["Special: Reverse Shadow Star System:", analysisStats.reverseShadowSystem || ' - ']
        ];

        autoTable(doc, {
          startY: currentY,
          margin: { left: 15, right: 15 },
          body: basicFactsBody,
          theme: 'grid',
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 },
          columnStyles: {
            0: { cellWidth: 70, fontStyle: 'bold', fillColor: [248, 248, 248] },
            1: { cellWidth: 110 }
          }
        });
        
        // @ts-ignore
        currentY = doc.lastAutoTable.finalY + 12;

        // Demographics heading
        currentY = drawSubsectionHeading(doc, "Demographics Breakdown", currentY);

        // Render Demographics Table on the right and Demographics Pie Chart on the left
        const factionsPieImg = generatePieChartDataUrl(analysisStats.factionsPieData);
        if (factionsPieImg) {
          doc.addImage(factionsPieImg, 'PNG', 15, currentY, 40, 40);
        }

        const factionsTableBody = analysisStats.factionsPieData.map(item => {
          const totalVal = analysisStats.factionsPieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });

        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Faction Label", "Systems", "Proportion"]],
          body: factionsTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2.5, lineColor: [220, 220, 200], lineWidth: 0.1 },
          columnStyles: {
            0: { fontStyle: 'bold' },
            1: { halign: 'right' },
            2: { halign: 'right' }
          }
        });

        // --- REPORT PAGE 2 (PDF Page 3) ---
        createNewPage(2);
        currentY = 30;

        currentY = drawSectionHeading(doc, "SYSTEM PARAMETERS DISTRIBUTIONS", currentY);

        // Star Color Breakdown (Pie left, Table right)
        currentY = drawSubsectionHeading(doc, "Star Color Breakdown", currentY);
        const starColorPieImg = generatePieChartDataUrl(analysisStats.starColorPieData);
        if (starColorPieImg) {
          doc.addImage(starColorPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const starColorTableBody = analysisStats.starColorPieData.map(item => {
          const totalVal = analysisStats.starColorPieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Star Color", "Systems", "Proportion"]],
          body: starColorTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });
        // @ts-ignore
        currentY = doc.lastAutoTable.finalY + 12;

        // Dissonance Breakdown (Pie left, Table right)
        currentY = drawSubsectionHeading(doc, "Dissonance Breakdown", currentY);
        const dissonancePieImg = generatePieChartDataUrl(analysisStats.dissonancePieData);
        if (dissonancePieImg) {
          doc.addImage(dissonancePieImg, 'PNG', 15, currentY, 40, 40);
        }
        const dissonanceTableBody = analysisStats.dissonancePieData.map(item => {
          const totalVal = analysisStats.dissonancePieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Dissonance Level", "Systems", "Proportion"]],
          body: dissonanceTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 3 (PDF Page 4) ---
        createNewPage(3);
        currentY = 30;

        // Giant Systems Breakdown
        currentY = drawSubsectionHeading(doc, "Giant Systems Breakdown", currentY);
        const giantPieImg = generatePieChartDataUrl(analysisStats.giantPieData);
        if (giantPieImg) {
          doc.addImage(giantPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const giantTableBody = analysisStats.giantPieData.map(item => {
          const totalVal = analysisStats.giantPieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Giant Presence", "Systems", "Proportion"]],
          body: giantTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });
        // @ts-ignore
        currentY = doc.lastAutoTable.finalY + 12;

        // Water Presence Breakdown
        currentY = drawSubsectionHeading(doc, "Water Breakdown", currentY);
        const waterPieImg = generatePieChartDataUrl(analysisStats.waterPieData);
        if (waterPieImg) {
          doc.addImage(waterPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const waterTableBody = analysisStats.waterPieData.map(item => {
          const totalVal = analysisStats.waterPieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Water Presence", "Systems", "Proportion"]],
          body: waterTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 4 (PDF Page 5) ---
        createNewPage(4);
        currentY = 30;

        // Wealth Level Breakdown
        currentY = drawSubsectionHeading(doc, "Wealth Level Breakdown", currentY);
        const wealthPieImg = generatePieChartDataUrl(analysisStats.wealthPieData);
        if (wealthPieImg) {
          doc.addImage(wealthPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const wealthTableBody = analysisStats.wealthPieData.map(item => {
          const totalVal = analysisStats.wealthPieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Wealth Level", "Systems", "Proportion"]],
          body: wealthTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });
        // @ts-ignore
        currentY = doc.lastAutoTable.finalY + 12;

        // Economy Type Breakdown
        currentY = drawSubsectionHeading(doc, "Economy Type Breakdown", currentY);
        const economyPieImg = generatePieChartDataUrl(analysisStats.economyTypePieData);
        if (economyPieImg) {
          doc.addImage(economyPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const economyTableBody = analysisStats.economyTypePieData.map(item => {
          const totalVal = analysisStats.economyTypePieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Economy Type", "Systems", "Proportion"]],
          body: economyTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 5 (PDF Page 6) ---
        createNewPage(5);
        currentY = 30;

        // Conflict Level Breakdown
        currentY = drawSubsectionHeading(doc, "Conflict Level Breakdown", currentY);
        const conflictPieImg = generatePieChartDataUrl(analysisStats.conflictPieData);
        if (conflictPieImg) {
          doc.addImage(conflictPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const conflictTableBody = analysisStats.conflictPieData.map(item => {
          const totalVal = analysisStats.conflictPieData.reduce((acc, x) => acc + x.value, 0);
          const percent = totalVal > 0 ? Math.round((item.value / totalVal) * 100) : 0;
          return [item.label, `${item.value}`, `${percent}%`];
        });
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Conflict Level", "Systems", "Proportion"]],
          body: conflictTableBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 6 (PDF Page 7) ---
        createNewPage(6);
        currentY = 30;

        currentY = drawSectionHeading(doc, "SECTION 2: SYSTEM CLAIMING", currentY);

        currentY = drawSubsectionHeading(doc, "Civilization Claims Ledger Summary", currentY);
        const claimsPieImg = generatePieChartDataUrl(analysisStats.systemClaimsPieData);
        if (claimsPieImg) {
          doc.addImage(claimsPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const claimsTableBodyMapped = analysisStats.claimsTableData.map(item => [
          item.civ,
          `${item.count}`,
          `${item.percent}%`
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Civilization Label", "Systems", "Proportion"]],
          body: claimsTableBodyMapped,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 7 (PDF Page 8) ---
        createNewPage(7);
        currentY = 30;

        currentY = drawSectionHeading(doc, "SECTION 3: DISCOVERY AND SURVEYORS", currentY);

        currentY = drawSubsectionHeading(doc, "Discoverer Distribution", currentY);
        const discoverersPieImg = generatePieChartDataUrl(analysisStats.discoverersPieData);
        if (discoverersPieImg) {
          doc.addImage(discoverersPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const topDiscoverersBody = analysisStats.top10Discoverers.map(item => [
          item.name,
          `${item.count}`
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Discoverer Name", "Systems Scanned"]],
          body: topDiscoverersBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 8 (PDF Page 9) ---
        createNewPage(8);
        currentY = 30;

        currentY = drawSubsectionHeading(doc, "Surveyor Distribution", currentY);
        const surveyorsPieImg = generatePieChartDataUrl(analysisStats.surveyorsPieData);
        if (surveyorsPieImg) {
          doc.addImage(surveyorsPieImg, 'PNG', 15, currentY, 40, 40);
        }
        const topSurveyorsBody = analysisStats.top10Surveyors.map(item => [
          item.name,
          `${item.count}`
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 60, right: 15 },
          head: [["Surveyor Name", "Systems Surveyed"]],
          body: topSurveyorsBody,
          theme: 'grid',
          headStyles: { fillColor: [230, 230, 230], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 8, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 9 (PDF Page 10) ---
        createNewPage(9);
        currentY = 30;

        currentY = drawSubsectionHeading(doc, "Registry Temporal Ledgers (Discoveries)", currentY);
        
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(9);
        doc.setTextColor(0, 0, 0);
        doc.text("10 Earliest Discovery Dates", 15, currentY);
        currentY += 4;

        const earliestDiscBody = analysisStats.earliestDiscoveryDates.map(item => [
          item.system,
          item.dateStr,
          item.name
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 15, right: 15 },
          head: [["System", "Date", "Discoverer"]],
          body: earliestDiscBody,
          theme: 'grid',
          headStyles: { fillColor: [235, 235, 235], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 7.5, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });
        
        // @ts-ignore
        currentY = doc.lastAutoTable.finalY + 8;

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(9);
        doc.text("10 Most Recent Discovery Dates", 15, currentY);
        currentY += 4;

        const recentDiscBody = analysisStats.mostRecentDiscoveryDates.map(item => [
          item.system,
          item.dateStr,
          item.name
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 15, right: 15 },
          head: [["System", "Date", "Discoverer"]],
          body: recentDiscBody,
          theme: 'grid',
          headStyles: { fillColor: [235, 235, 235], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 7.5, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 10 (PDF Page 11) ---
        createNewPage(10);
        currentY = 30;

        currentY = drawSubsectionHeading(doc, "Registry Temporal Ledgers (Surveyors)", currentY);
        
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(9);
        doc.text("10 Earliest Survey Dates", 15, currentY);
        currentY += 4;

        const earliestSurvBody = analysisStats.earliestSurveyorDates.map(item => [
          item.system,
          item.dateStr,
          item.name
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 15, right: 15 },
          head: [["System", "Date", "Surveyor"]],
          body: earliestSurvBody,
          theme: 'grid',
          headStyles: { fillColor: [235, 235, 235], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 7.5, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // @ts-ignore
        currentY = doc.lastAutoTable.finalY + 8;

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(9);
        doc.text("10 Most Recent Survey Dates", 15, currentY);
        currentY += 4;

        const recentSurvBody = analysisStats.mostRecentSurveyorDates.map(item => [
          item.system,
          item.dateStr,
          item.name
        ]);
        autoTable(doc, {
          startY: currentY,
          margin: { left: 15, right: 15 },
          head: [["System", "Date", "Surveyor"]],
          body: recentSurvBody,
          theme: 'grid',
          headStyles: { fillColor: [235, 235, 235], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 8 },
          styles: { textColor: [0, 0, 0], fontSize: 7.5, cellPadding: 2, lineColor: [220, 220, 220], lineWidth: 0.1 }
        });

        // --- REPORT PAGE 11 (PDF Page 12) ---
        createNewPage(11);
        currentY = 30;

        currentY = drawSectionHeading(doc, "SECTION 4: ADDITIONAL INFORMATION", currentY);

        // Center box for "For Future use" with white bg and grey border
        doc.setFont("Helvetica", "normal");
        doc.setFontSize(9);
        doc.setTextColor(120, 120, 120);
        doc.setDrawColor(220, 220, 220);
        doc.setFillColor(252, 252, 252);
        
        doc.rect(15, currentY + 5, 180, 40, 'FD');
        doc.text("For Future use", 105, currentY + 27, { align: 'center' });

        doc.save(`AGT_Region_Analysis_Report_${selectedRegion}.pdf`);
      } catch (err) {
        console.error("Error generating region analysis PDF report:", err);
      } finally {
        setReportGeneratingLoading(false);
      }
    }, 150);
  };

  const downloadFullReportPdf = () => {
    if (sortedMatchedRecords.length === 0) return;
    
    const activeCols = columns.filter(col => col.enabled);
    
    // Check if any row in the PDF will wrap into more than 3 lines.
    // Spec: "If it is not possible to generate a PDF where all individual rows will now fit in 3 lines or less, then display pop up an error box..."
    let exceedsThreeLines = false;
    
    if (activeCols.length > 13) {
      exceedsThreeLines = true;
    } else {
      for (const record of sortedMatchedRecords) {
        for (const col of activeCols) {
          const val = getDisplayValue(record[col.name], col.rawIndex);
          let colWidth = 257 / activeCols.length;
          
          if (col.rawIndex === 76) colWidth = 12;
          else if (col.rawIndex === 7) colWidth = 24;
          else if (col.rawIndex === 1) colWidth = 35;
          else if (col.rawIndex === 14 || col.rawIndex === 15) colWidth = 18;
          
          const charLimit = Math.max(5, Math.floor(colWidth / 1.5));
          const lines = Math.ceil((val || '').length / charLimit) || 1;
          if (lines > 3) {
            exceedsThreeLines = true;
            break;
          }
        }
        if (exceedsThreeLines) break;
      }
    }
    
    if (exceedsThreeLines) {
      setPdfErrorMsg("Too many columns for PDF report, reduce the column selections or choose CSV export");
      return;
    }

    setReportGeneratingLoading(true);

    setTimeout(() => {
      try {
        const doc = new jsPDF('l', 'mm', 'a4'); // Landscape for tables

        // --- 1. COVER PAGE (Page 1) ---
        // Title in hex color FF0500 (RGB: 255, 5, 0), horizontally centered under AGTIcon
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(26);
        doc.setTextColor(255, 5, 0);
        doc.text("AGT System Report", 148.5, 87, { align: 'center' });

        // All other text on the cover page is black
        doc.setFont("Helvetica", "normal");
        doc.setTextColor(0, 0, 0);
        doc.setFontSize(14);
        
        const formattedReportType = reportType === 'simple' ? 'Simple' : reportType === 'detailed' ? 'Detailed' : 'Custom';
        doc.text(`Report Type: ${formattedReportType}`, 148.5, 102, { align: 'center' });

        // Civilization translation rule
        let civDisplay = searchKey || 'All';
        if (civDisplay.trim().toLowerCase() === 'agt') {
          civDisplay = 'Alliance of Galactic Travellers';
        }

        doc.setFontSize(11);
        doc.text(`Civilization: ${civDisplay}`, 148.5, 115, { align: 'center' });
        
        // Combined Galaxy & Region filter criteria line
        const galaxyVal = selectedGalaxy || 'All';
        const regionVal = selectedRegion || 'All';
        doc.text(`Galaxy: ${galaxyVal} / Region: ${regionVal}`, 148.5, 122, { align: 'center' });

        // Discoverer & Surveyor line
        const discVal = discovererName.trim() || 'All';
        const survVal = surveyorName.trim() || 'All';
        doc.text(`Discoverer: ${discVal} / Surveyor: ${survVal}`, 148.5, 129, { align: 'center' });

        // Date of Report line
        doc.text(`Date of Report: ${new Date().toLocaleString()}`, 148.5, 138, { align: 'center' });

        // Result Count line
        doc.text(`Result Count: ${matchedRecords.length} Verified Entries`, 148.5, 145, { align: 'center' });

        // Top logo center image placement
        try {
          const imgElements = document.getElementsByTagName('img');
          let logoImg: HTMLImageElement | null = null;
          for (let i = 0; i < imgElements.length; i++) {
            if (imgElements[i].alt === 'AGT Logo') {
              logoImg = imgElements[i];
              break;
            }
          }
          if (logoImg && logoImg.complete) {
            doc.addImage(logoImg, 'PNG', 133.5, 42, 30, 30);
          } else {
            const tempImg = new Image();
            tempImg.src = logoUrl;
            if (tempImg.complete) {
              doc.addImage(tempImg, 'PNG', 133.5, 42, 30, 30);
            }
          }
        } catch (err) {
          console.warn('Cover page logo inject failed:', err);
        }

        // Star Systems Icon centering at bottom segment midpoint y=162.5
        try {
          const imgElements = document.getElementsByTagName('img');
          let starImg: HTMLImageElement | null = null;
          for (let i = 0; i < imgElements.length; i++) {
            if (imgElements[i].alt === 'Star Systems') {
              starImg = imgElements[i];
              break;
            }
          }
          if (starImg && starImg.complete) {
            doc.addImage(starImg, 'PNG', 133.5, 162.5, 30, 30);
          } else {
            const tempImg = new Image();
            tempImg.src = starSystemsIcon;
            if (tempImg.complete) {
              doc.addImage(tempImg, 'PNG', 133.5, 162.5, 30, 30);
            }
          }
        } catch (err) {
          console.warn('Cover page star logo inject failed:', err);
        }

        // --- 2. ADD PAGE 2 (TABLE DATA) ---
        doc.addPage();

        const urlMap = new Map<string, string>();
        const tableData = sortedMatchedRecords.map((record, rIdx) => 
          activeCols.map((col, cIdx) => {
            const rawVal = record[col.name];
            const val = getDisplayValue(rawVal, col.rawIndex);
            
            const isUrl = (col.rawIndex >= 76 && col.rawIndex <= 80) || String(rawVal || '').trim().startsWith('http');
            if (isUrl && rawVal) {
              urlMap.set(`${rIdx}-${cIdx}`, String(rawVal).trim());
              return 'LINK';
            }
            return val || '-';
          })
        );

        // Add total row to PDF
        const countFieldName = columns[0]?.name;
        const totalRow = activeCols.map(col => {
          if (col.name === countFieldName) return `Count: ${sortedMatchedRecords.length}`;
          return '';
        });
        tableData.push(totalRow);

        // Compute Column Width overrides subject to rules:
        // - Wiki (Col BY): slim width (10-12mm)
        // - Region (Col B - Col 1): maximum 2 lines (custom styling width: 25mm)
        // - Coordinate: maximum 1 line (width: 20mm, ellipsis overflow)
        // - Discovery Date (index 14) and Surveyor Date (index 15): fit 1 line (width: 16mm)
        // AND ensure the last column text field is not truncated in the produced reports, adjusting other column widths.
        const lastColIdx = activeCols.length - 1;
        const columnStyles: Record<number, any> = {};
        
        let allocatedWidth = 0;
        
        activeCols.forEach((col, idx) => {
          if (idx === lastColIdx) {
            // Last column: will be set precisely to remaining width and use linebreak overflow
            columnStyles[idx] = { cellWidth: 'auto', overflow: 'linebreak' };
          } else {
            if (col.rawIndex === 76) {
              columnStyles[idx] = { cellWidth: 10, overflow: 'ellipsize' };
              allocatedWidth += 10;
            } else if (col.rawIndex === 7) {
              columnStyles[idx] = { cellWidth: 20, overflow: 'ellipsize' };
              allocatedWidth += 20;
            } else if (col.rawIndex === 1) {
              columnStyles[idx] = { cellWidth: 25, overflow: 'ellipsize' };
              allocatedWidth += 25;
            } else if (col.rawIndex === 14 || col.rawIndex === 15) {
              columnStyles[idx] = { cellWidth: 16, overflow: 'ellipsize' };
              allocatedWidth += 16;
            } else {
              const qtyOtherStandardCols = activeCols.filter((c, i) => 
                i !== lastColIdx && c.rawIndex !== 76 && c.rawIndex !== 7 && c.rawIndex !== 1 && c.rawIndex !== 14 && c.rawIndex !== 15
              ).length || 1;
              
              const totalForOtherStandards = 257 - allocatedWidth - 45 - (activeCols.filter((c, i) => i !== lastColIdx && (c.rawIndex === 76 || c.rawIndex === 7 || c.rawIndex === 1 || c.rawIndex === 14 || c.rawIndex === 15)).length * 15);
              const computedStandardColWidth = Math.max(12, Math.floor(totalForOtherStandards / qtyOtherStandardCols));
              columnStyles[idx] = { cellWidth: computedStandardColWidth, overflow: 'ellipsize' };
            }
          }
        });
        
        // Resolve reserved last column width beautifully
        const reservedLastWidth = Math.max(45, 257 - activeCols.reduce((sum, col, idx) => {
          if (idx === lastColIdx) return sum;
          const cw = columnStyles[idx]?.cellWidth;
          return sum + (typeof cw === 'number' ? cw : 15);
        }, 0));
        
        columnStyles[lastColIdx] = { cellWidth: reservedLastWidth, overflow: 'linebreak' };

        autoTable(doc, {
          startY: 25,
          head: [activeCols.map(col => {
            if (col.rawIndex === 76) return "Wiki";
            if (col.rawIndex === 10) return "Surveyor";
            return col.name;
          })],
          body: tableData,
          theme: 'grid',
          columnStyles: columnStyles,
          headStyles: { fillColor: [20, 20, 20], textColor: [255, 255, 255], fontSize: 8 },
          bodyStyles: { fontSize: 8 },
          footStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold' },
          margin: { top: 25, bottom: 20, left: 20, right: 20 },
          didParseCell: (dataCell) => {
            if (dataCell.row.index === tableData.length - 1) {
              dataCell.cell.styles.fillColor = [220, 220, 220];
              dataCell.cell.styles.fontStyle = 'bold';
            }
            
            const cellKey = `${dataCell.row.index}-${dataCell.column.index}`;
            if (urlMap.has(cellKey)) {
              dataCell.cell.styles.textColor = [0, 0, 255];
            }
          },
          didDrawCell: (dataCell) => {
            const cellKey = `${dataCell.row.index}-${dataCell.column.index}`;
            const url = urlMap.get(cellKey);
            if (url && dataCell.section === 'body') {
              doc.link(dataCell.cell.x, dataCell.cell.y, dataCell.cell.width, dataCell.cell.height, { url });
            }
          }
        });

        // --- 3. POST-PROCESS PAGES FOR HEADERS & FOOTERS (Page 2 onwards) ---
        const totalPagesCount = doc.getNumberOfPages();
        for (let i = 2; i <= totalPagesCount; i++) {
          doc.setPage(i);
          
          // Header left justified: Small logo (10x10) and text
          try {
            const imgElements = document.getElementsByTagName('img');
            let logoImg: HTMLImageElement | null = null;
            for (let j = 0; j < imgElements.length; j++) {
              if (imgElements[j].alt === 'AGT Logo') {
                logoImg = imgElements[j];
                break;
              }
            }
            if (logoImg && logoImg.complete) {
              doc.addImage(logoImg, 'PNG', 20, 10, 10, 10);
            } else {
              const tempImg = new Image();
              tempImg.src = logoUrl;
              if (tempImg.complete) {
                doc.addImage(tempImg, 'PNG', 20, 10, 10, 10);
              }
            }
          } catch (err) {
            console.warn('Page header logo inject failed:', err);
          }

          doc.setFont("Helvetica", "bold");
          doc.setFontSize(10);
          doc.setTextColor(0, 0, 0);
          doc.text("AGT System Report", 32, 16.5);

          // Header right justified: Page number starting at 1 following cover page
          doc.setFont("Helvetica", "normal");
          doc.setFontSize(9);
          doc.text(`Page ${i - 1}`, 277, 16.5, { align: 'right' });

          // Footer left justified: Date of Report
          doc.setFont("Helvetica", "normal");
          doc.setFontSize(8);
          doc.text(`Date of Report: ${new Date().toLocaleString()}`, 20, 202);
        }

        doc.save(`AGT  System Report - ${formatTimestamp()}.pdf`);
      } catch (err) {
        console.error('Error generating PDF:', err);
      } finally {
        setReportGeneratingLoading(false);
      }
    }, 400);
  };

  const downloadCsv = () => {
    if (sortedMatchedRecords.length === 0) return;
    setReportGeneratingLoading(true);

    setTimeout(() => {
      try {
        const activeCols = columns.filter(col => col.enabled);
        const csvData = sortedMatchedRecords.map(record => {
          const row: any = {};
          activeCols.forEach(col => {
            row[col.name] = getDisplayValue(record[col.name], col.rawIndex);
          });
          return row;
        });
        
        const csv = Papa.unparse(csvData);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        
        if (link.download !== undefined) {
          const url = URL.createObjectURL(blob);
          link.setAttribute('href', url);
          link.setAttribute('download', `AGT  System CSV Export- ${formatTimestamp()}.csv`);
          link.style.visibility = 'hidden';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
        }
      } catch (err) {
        console.error('Error generating CSV:', err);
      } finally {
        setReportGeneratingLoading(false);
      }
    }, 400);
  };

  const toggleColumn = (name: string) => {
    setColumns(prev => prev.map(c => c.name === name ? { ...c, enabled: !c.enabled } : c));
  };

  const activeColumnsCount = useMemo(() => columns.filter(c => c.enabled).length, [columns]);

  const sortedMatchedRecords = useMemo(() => {
    if (!sortColumn || !sortDirection) return matchedRecords;
    
    return [...matchedRecords].sort((a, b) => {
      const valA = a[sortColumn];
      const valB = b[sortColumn];
      
      if (valA === undefined || valB === undefined) return 0;
      
      const strA = String(valA || '').trim();
      const strB = String(valB || '').trim();
      
      // Try numeric sort
      const cleanA = strA.replace(/[$,%]/g, '');
      const cleanB = strB.replace(/[$,%]/g, '');
      const numA = parseFloat(cleanA);
      const numB = parseFloat(cleanB);
      
      const isNumA = !isNaN(numA) && isFinite(numA) && cleanA !== '';
      const isNumB = !isNaN(numB) && isFinite(numB) && cleanB !== '';
      
      if (isNumA && isNumB) {
        return sortDirection === 'asc' ? numA - numB : numB - numA;
      }
      
      return sortDirection === 'asc'
        ? strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' })
        : strB.localeCompare(strA, undefined, { numeric: true, sensitivity: 'base' });
    });
  }, [matchedRecords, sortColumn, sortDirection]);

  const totalPoints = useMemo(() => {
    return sortedMatchedRecords.length;
  }, [sortedMatchedRecords]);

  const totalPages = useMemo(() => {
    return Math.ceil(sortedMatchedRecords.length / pageSize);
  }, [sortedMatchedRecords.length, pageSize]);

  const paginatedRecords = useMemo(() => {
    const startIdx = (currentPage - 1) * pageSize;
    return sortedMatchedRecords.slice(startIdx, startIdx + pageSize);
  }, [sortedMatchedRecords, currentPage, pageSize]);

  const multiplier = parseFloat(fontScale.replace('x', '')) || 1.0;

  const analysisStats = useMemo(() => {
    if (matchedRecords.length === 0) return null;
    
    // basic facts
    const firstRow = matchedRecords[0];
    const regionAge = getCellValue(firstRow, 'CL') || 'N/A';
    
    // claims status: check if any S is not blank
    const systemsWithClaims = matchedRecords.filter(r => getCellValue(r, 'S') !== '');
    const regionClaimStatus = systemsWithClaims.length > 0 
      ? `${systemsWithClaims.length}/${matchedRecords.length}`
      : 'Unclaimed';
      
    // base address
    const firstCoordRow = matchedRecords.find(r => getCellValue(r, 'H'));
    const coordinateBaseAddress = firstCoordRow 
      ? getCellValue(firstCoordRow, 'H').substring(0, 14) 
      : 'N/A';
      
    // galaxy
    const galaxyName = getCellValue(firstRow, 'A') || 'N/A';
    
    // Special Systems
    // Black Hole (Q matches "Black Hole")
    const bhRow = matchedRecords.find(r => getCellValue(r, 'Q').toLowerCase() === 'black hole');
    const blackHoleSystem = bhRow ? getCellValue(bhRow, 'C') : ' - ';
    
    // Atlas (Q matches "Atlas")
    const atlasRow = matchedRecords.find(r => getCellValue(r, 'Q').toLowerCase() === 'atlas');
    const atlasSystem = atlasRow ? getCellValue(atlasRow, 'C') : ' - ';
    
    // Shadow Star (BV matches "Shadow")
    const shadowRow = matchedRecords.find(r => getCellValue(r, 'BV').toLowerCase() === 'shadow');
    const shadowSystem = shadowRow ? getCellValue(shadowRow, 'C') : ' - ';
    
    let estimatedNormalSystems = 'Unknown';
    if (shadowRow) {
      const coords = getCellValue(shadowRow, 'H').trim();
      const hexOnly = coords.replace(/[^0-9A-Fa-f]/g, '');
      if (hexOnly.length >= 3) {
        const lastThree = hexOnly.slice(-3);
        const decimalVal = parseInt(lastThree, 16);
        if (!isNaN(decimalVal)) {
          estimatedNormalSystems = String(decimalVal - 1);
        }
      }
    }
    
    // Reverse Shadow Star (BV matches "Reverse Shadow")
    const revShadowRow = matchedRecords.find(r => getCellValue(r, 'BV').toLowerCase() === 'reverse shadow');
    const reverseShadowSystem = revShadowRow ? getCellValue(revShadowRow, 'C') : ' - ';
    
    // Factions Pie Data (AC)
    const factionCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const f = getCellValue(r, 'AC').trim();
      let label = 'Unknown';
      if (f === '' || f.toLowerCase() === 'none' || f.toLowerCase() === 'null') {
        label = 'None';
      } else {
        label = f;
      }
      factionCounts[label] = (factionCounts[label] || 0) + 1;
    });
    const defaultColors = [
      '#FFB451', // AGT Orange
      '#EF4444', // Red
      '#3B82F6', // Blue
      '#10B981', // Emerald Green
      '#8B5CF6', // Purple
      '#F59E0B', // Amber
      '#06B6D4', // Cyan
      '#EC4899', // Pink
      '#14B8A6', // Teal
      '#6366F1'  // Indigo
    ];
    const factionsPieData = Object.entries(factionCounts).map(([label, value], index) => ({
      label,
      value,
      color: defaultColors[index % defaultColors.length]
    }));
    
    // Star Color Pie Data (Z)
    const starColorCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const col = getCellValue(r, 'Z').trim();
      const label = (col === '' || col.toLowerCase() === 'n/a') ? 'Unknown' : col;
      starColorCounts[label] = (starColorCounts[label] || 0) + 1;
    });
    const colorMap: Record<string, string> = {
      yellow: '#FBBF24',
      red: '#EF4444',
      blue: '#3B82F6',
      green: '#10B981',
      white: '#FFFFFF',
      orange: '#F97316',
      teal: '#06B6D4'
    };
    const starColorPieData = Object.entries(starColorCounts).map(([label, value]) => {
      const lowerTheme = label.toLowerCase().trim();
      let color = colorMap[lowerTheme] || '#A855F7';
      if (lowerTheme === 'black hole' || lowerTheme === 'unknown') {
        color = '#FFFFFF';
      }
      return {
        label,
        value,
        color
      };
    });
    
    // Dissonance Pie Data (R)
    const dissonanceCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      let val = getCellValue(r, 'R').trim().toUpperCase();
      let label = val;
      if (val === 'Y') {
        label = 'Yes';
      } else if (val === 'N') {
        label = 'No';
      } else if (val === '') {
        label = 'Unknown';
      }
      dissonanceCounts[label] = (dissonanceCounts[label] || 0) + 1;
    });
    const dissonancePieData = Object.entries(dissonanceCounts).map(([label, value]) => {
      let color = '#A855F7';
      if (label === 'Yes') color = '#EF4444';
      if (label === 'No') color = '#10B981';
      if (label === 'Unknown') color = '#6B7280';
      return { label, value, color };
    });
    
    // Giant Systems Pie Data (N)
    const giantCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const val = getCellValue(r, 'N').trim().toUpperCase();
      let label = 'Unknown';
      if (val === 'Y') label = 'Yes';
      else if (val === 'N') label = 'No';
      giantCounts[label] = (giantCounts[label] || 0) + 1;
    });
    const giantPieData = Object.entries(giantCounts).map(([label, value]) => {
      let color = '#FFB451';
      if (label === 'No') color = '#3B82F6';
      if (label === 'Yes') color = '#EF4444';
      if (label === 'Unknown') color = '#6B7280';
      return { label, value, color };
    });
    
    // Water Presence (AF)
    const waterCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const val = getCellValue(r, 'AF').trim().toUpperCase();
      let label = 'Unknown';
      if (val === 'Y') label = 'Yes';
      else if (val === 'N') label = 'No';
      waterCounts[label] = (waterCounts[label] || 0) + 1;
    });
    const waterPieData = Object.entries(waterCounts).map(([label, value]) => {
      let color = '#3B82F6';
      if (label === 'No') color = '#F59E0B';
      if (label === 'Unknown') color = '#6B7280';
      return { label, value, color };
    });
    
    // --- Economics and Conflict ---
    // Wealth Level Pie (CY) - non-blank only
    const wealthCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const val = getCellValue(r, 'CY');
      if (val) {
        wealthCounts[val] = (wealthCounts[val] || 0) + 1;
      }
    });
    const wealthLevelMap: Record<string, { label: string; color: string }> = {
      '0': { label: 'Unknown', color: '#6B7280' },
      'unknown': { label: 'Unknown', color: '#6B7280' },
      '1': { label: '1 - Poor', color: '#EF4444' },
      'poor': { label: '1 - Poor', color: '#EF4444' },
      '2': { label: '2 - Moderate', color: '#FBBF24' },
      'moderate': { label: '2 - Moderate', color: '#FBBF24' },
      '3': { label: '3 - Wealthy', color: '#10B981' },
      'wealthy': { label: '3 - Wealthy', color: '#10B981' }
    };
    const wealthPieData = Object.entries(wealthCounts).map(([key, value]) => {
      const mapped = wealthLevelMap[key.toLowerCase()] || { label: key, color: '#A855F7' };
      return { label: mapped.label, value, color: mapped.color };
    });
    
    // Economy Type Pie (CX) - non-blank only
    const colorsPalette = [
      '#FFB451', // AGT Orange
      '#EF4444', // Red
      '#3B82F6', // Blue
      '#10B981', // Emerald Green
      '#8B5CF6', // Purple
      '#F59E0B', // Amber
      '#06B6D4', // Cyan
      '#EC4899', // Pink
      '#14B8A6', // Teal
      '#6366F1'  // Indigo
    ];
    const economyTypeCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const val = getCellValue(r, 'CX');
      if (val) {
        economyTypeCounts[val] = (economyTypeCounts[val] || 0) + 1;
      }
    });
    const economyTypePieData = Object.entries(economyTypeCounts).map(([label, value], index) => ({
      label,
      value,
      color: colorsPalette[index % colorsPalette.length]
    }));
    
    // Conflict Pie (CZ) - non-blank only
    const conflictCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const val = getCellValue(r, 'CZ');
      if (val) {
        conflictCounts[val] = (conflictCounts[val] || 0) + 1;
      }
    });
    const conflictLevelMap: Record<string, { label: string; color: string }> = {
      '0': { label: 'Unknown', color: '#6B7280' },
      'unknown': { label: 'Unknown', color: '#6B7280' },
      '1': { label: '1 - Low', color: '#10B981' },
      'low': { label: '1 - Low', color: '#10B981' },
      '2': { label: '2 - Moderate', color: '#FBBF24' },
      'moderate': { label: '2 - Moderate', color: '#FBBF24' },
      '3': { label: '3 - Violent', color: '#EF4444' },
      'violent': { label: '3 - Violent', color: '#EF4444' }
    };
    const conflictPieData = Object.entries(conflictCounts).map(([key, value]) => {
      const mapped = conflictLevelMap[key.toLowerCase()] || { label: key, color: '#A855F7' };
      return { label: mapped.label, value, color: mapped.color };
    });
    
    // Notable Economies min/max values
    let maxBuyVal = -Infinity;
    let minBuyVal = Infinity;
    let maxSellVal = -Infinity;
    let minSellVal = Infinity;
    
    const recordsWithProducts = matchedRecords.map(r => {
      const buy = parseTradeValue(getCellValue(r, 'AI'));
      const sell = parseTradeValue(getCellValue(r, 'AJ'));
      return { r, buy, sell, name: getCellValue(r, 'C') };
    });
    
    recordsWithProducts.forEach(item => {
      const rawBuy = getCellValue(item.r, 'AI');
      const rawSell = getCellValue(item.r, 'AJ');
      if (rawBuy) {
        if (item.buy > maxBuyVal) maxBuyVal = item.buy;
        if (item.buy < minBuyVal) minBuyVal = item.buy;
      }
      if (rawSell) {
        if (item.sell > maxSellVal) maxSellVal = item.sell;
        if (item.sell < minSellVal) minSellVal = item.sell;
      }
    });
    
    const highestBuySystems = Array.from(new Set(recordsWithProducts.filter(item => getCellValue(item.r, 'AI') && item.buy === maxBuyVal).map(item => item.name))).join(', ');
    const lowestBuySystems = Array.from(new Set(recordsWithProducts.filter(item => getCellValue(item.r, 'AI') && item.buy === minBuyVal).map(item => item.name))).join(', ');
    const highestSellSystems = Array.from(new Set(recordsWithProducts.filter(item => getCellValue(item.r, 'AJ') && item.sell === maxSellVal).map(item => item.name))).join(', ');
    const lowestSellSystems = Array.from(new Set(recordsWithProducts.filter(item => getCellValue(item.r, 'AJ') && item.sell === minSellVal).map(item => item.name))).join(', ');
    
    // --- Claims count table (Section 2) ---
    const claimsProductCounts: Record<string, number> = {};
    matchedRecords.forEach(r => {
      const civ = getCellValue(r, 'S') || 'Unclaimed / Independent';
      claimsProductCounts[civ] = (claimsProductCounts[civ] || 0) + 1;
    });
    
    const claimsList = Object.entries(claimsProductCounts)
      .map(([civ, count]) => ({ civ, count }))
      .sort((a, b) => b.count - a.count);

    const claimsWithColors = claimsList.map((item, index) => {
      let color = colorsPalette[index % colorsPalette.length];
      if (item.civ === 'Alliance of Galactic Travellers') {
        color = '#FFB451';
      }
      return {
        ...item,
        color
      };
    });

    const claimsTableData = claimsWithColors.map(item => ({
      civ: item.civ,
      count: item.count,
      percent: matchedRecords.length > 0 ? Math.round((item.count / matchedRecords.length) * 100) : 0,
      color: item.color
    }));
      
    // System Claims Pie Data (S)
    const systemClaimsPieData = claimsWithColors.map(item => ({
      label: item.civ,
      value: item.count,
      color: item.color
    }));
    
    // --- Discoverers (Section 3) ---
    const discovererCountsLower: Record<string, number> = {};
    const discovererCasingMap: Record<string, string> = {}; 
    
    matchedRecords.forEach(r => {
      const disc = getCellValue(r, 'L');
      if (disc) {
        const discTrimmed = disc.trim();
        if (discTrimmed) {
          const discLower = discTrimmed.toLowerCase();
          discovererCountsLower[discLower] = (discovererCountsLower[discLower] || 0) + 1;
          if (!discovererCasingMap[discLower]) {
            discovererCasingMap[discLower] = discTrimmed;
          }
        }
      }
    });
    
    const discovererCounts: Record<string, number> = {};
    Object.entries(discovererCountsLower).forEach(([discLower, count]) => {
      const originalCasing = discovererCasingMap[discLower];
      discovererCounts[originalCasing] = count;
    });

    const top10Discoverers = Object.entries(discovererCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);
      
    let discoverersPieData = Object.entries(discovererCounts)
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value);
    if (discoverersPieData.length > 5) {
      const top5 = discoverersPieData.slice(0, 5);
      const othersVal = discoverersPieData.slice(5).reduce((sum, item) => sum + item.value, 0);
      top5.push({ label: 'Others', value: othersVal });
      discoverersPieData = top5;
    }
    
    // Discovery dates
    const discovererDatesList = matchedRecords
      .filter(r => getCellValue(r, 'O') && getCellValue(r, 'L'))
      .map(r => ({
        system: getCellValue(r, 'C'),
        dateStr: getCellValue(r, 'O'),
        date: parseDateSafe(getCellValue(r, 'O')),
        name: getCellValue(r, 'L')
      }));
      
    const earliestDiscoveryDates = [...discovererDatesList]
      .sort((a, b) => a.date.getTime() - b.date.getTime())
      .slice(0, 10);
      
    const mostRecentDiscoveryDates = [...discovererDatesList]
      .sort((a, b) => b.date.getTime() - a.date.getTime())
      .slice(0, 10);
      
    // --- Surveyors (Section 3) ---
    const surveyorCountsLower: Record<string, number> = {};
    const surveyorCasingMap: Record<string, string> = {}; 
    
    matchedRecords.forEach(r => {
      const surv = getCellValue(r, 'K');
      if (surv) {
        const survTrimmed = surv.trim();
        if (survTrimmed) {
          const survLower = survTrimmed.toLowerCase();
          surveyorCountsLower[survLower] = (surveyorCountsLower[survLower] || 0) + 1;
          if (!surveyorCasingMap[survLower]) {
            surveyorCasingMap[survLower] = survTrimmed;
          }
        }
      }
    });
    
    const surveyorCounts: Record<string, number> = {};
    Object.entries(surveyorCountsLower).forEach(([survLower, count]) => {
      const originalCasing = surveyorCasingMap[survLower];
      surveyorCounts[originalCasing] = count;
    });

    const top10Surveyors = Object.entries(surveyorCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);
      
    let surveyorsPieData = Object.entries(surveyorCounts)
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value);
    if (surveyorsPieData.length > 5) {
      const top5 = surveyorsPieData.slice(0, 5);
      const othersVal = surveyorsPieData.slice(5).reduce((sum, item) => sum + item.value, 0);
      top5.push({ label: 'Others', value: othersVal });
      surveyorsPieData = top5;
    }
    
    // Surveyor Dates (column P along with K)
    const surveyorDatesList = matchedRecords
      .filter(r => getCellValue(r, 'P') && getCellValue(r, 'K'))
      .map(r => ({
        system: getCellValue(r, 'C'),
        dateStr: getCellValue(r, 'P'),
        date: parseDateSafe(getCellValue(r, 'P')),
        name: getCellValue(r, 'K')
      }));
      
    const earliestSurveyorDates = [...surveyorDatesList]
      .sort((a, b) => a.date.getTime() - b.date.getTime())
      .slice(0, 10);
      
    const mostRecentSurveyorDates = [...surveyorDatesList]
      .sort((a, b) => b.date.getTime() - a.date.getTime())
      .slice(0, 10);
      
    return {
      regionAge,
      regionClaimStatus,
      coordinateBaseAddress,
      galaxyName,
      blackHoleSystem,
      atlasSystem,
      shadowSystem,
      reverseShadowSystem,
      estimatedNormalSystems,
      factionsPieData,
      starColorPieData,
      dissonancePieData,
      giantPieData,
      waterPieData,
      wealthPieData,
      economyTypePieData,
      conflictPieData,
      highestBuy: highestBuySystems || ' - ',
      lowestBuy: lowestBuySystems || ' - ',
      highestSell: highestSellSystems || ' - ',
      lowestSell: lowestSellSystems || ' - ',
      highestBuyVal: maxBuyVal !== -Infinity ? `${maxBuyVal}%` : '',
      lowestBuyVal: minBuyVal !== Infinity ? `${minBuyVal}%` : '',
      highestSellVal: maxSellVal !== -Infinity ? `${maxSellVal}%` : '',
      lowestSellVal: minSellVal !== Infinity ? `${minSellVal}%` : '',
      claimsTableData,
      systemClaimsPieData,
      top10Discoverers,
      discoverersPieData,
      earliestDiscoveryDates,
      mostRecentDiscoveryDates,
      top10Surveyors,
      surveyorsPieData,
      earliestSurveyorDates,
      mostRecentSurveyorDates
    };
  }, [matchedRecords]);


  return (
    <div 
      onMouseDown={handleManualPlay}
      onTouchStart={handleManualPlay}
      className={`min-h-screen bg-[#0a0a0a] text-agt-orange font-sans selection:bg-agt-orange selection:text-black ${fontScale !== '1x' ? 'font-scale-active' : ''}`}
    >
      {/* Dynamic Font Scale Injector */}
      <style>{`
        .font-scale-active .text-\\[7px\\] { font-size: calc(7px * ${multiplier}) !important; }
        .font-scale-active .text-\\[8px\\] { font-size: calc(8px * ${multiplier}) !important; }
        .font-scale-active .text-\\[9px\\] { font-size: calc(9px * ${multiplier}) !important; }
        .font-scale-active .text-\\[10px\\] { font-size: calc(10px * ${multiplier}) !important; }
        .font-scale-active .text-\\[11px\\] { font-size: calc(11px * ${multiplier}) !important; }
        .font-scale-active .text-xs { font-size: calc(12px * ${multiplier}) !important; }
        .font-scale-active .text-sm { font-size: calc(14px * ${multiplier}) !important; }
        .font-scale-active .text-base { font-size: calc(16px * ${multiplier}) !important; }
        .font-scale-active .text-md { font-size: calc(16px * ${multiplier}) !important; }
        .font-scale-active .text-lg { font-size: calc(18px * ${multiplier}) !important; }
        .font-scale-active .text-xl { font-size: calc(20px * ${multiplier}) !important; }
        .font-scale-active .text-2xl { font-size: calc(24px * ${multiplier}) !important; }
        .font-scale-active .text-3xl { font-size: calc(30px * ${multiplier}) !important; }
        .font-scale-active .text-4xl { font-size: calc(36px * ${multiplier}) !important; }
        .font-scale-active .text-5xl { font-size: calc(48px * ${multiplier}) !important; }
        .font-scale-active .text-6xl { font-size: calc(60px * ${multiplier}) !important; }
        
        .font-scale-active input, 
        .font-scale-active select, 
        .font-scale-active textarea, 
        .font-scale-active button { 
          font-size: calc(100% * ${multiplier}) !important; 
        }
      `}</style>

      {/* Horizontally spinning logo processing overlay */}
      <AnimatePresence>
        {(searchLoading || reportGeneratingLoading) && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[200] bg-black/95 flex flex-col items-center justify-center p-6 text-center select-none"
          >
            <div className="relative w-48 h-48 flex items-center justify-center mb-8">
              {/* Outer orbit circle */}
              <div className="absolute inset-0 rounded-full border border-[#FF0500]/20 animate-pulse"></div>
              
              {/* Horizontally rotating logo img */}
              <motion.img 
                src="/AGTIcon.png" 
                alt="AGT Logo" 
                className="w-36 h-36 object-contain"
                animate={{ rotateY: 360 }}
                transition={{ repeat: Infinity, duration: 2, ease: "linear" }}
              />
            </div>
            <h3 className="text-[#FFB451] text-lg font-bold uppercase tracking-[0.25em] mb-2 animate-pulse">
              {reportGeneratingLoading ? t("Creating Report") : t("Searching AGT Galactic Archives")}
            </h3>
            <p className="text-[#FFB451]/50 text-xs font-mono uppercase tracking-[0.2em]">
              {reportGeneratingLoading ? t("Compiling intelligence packet...") : t("Establishing high-speed subspace connection...")}
            </p>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header */}
      <header className="border-b border-agt-orange/5 bg-black/40 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-5xl mx-auto px-6 h-20 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="relative w-10 h-10 flex items-center justify-center shrink-0">
              <img 
                src={logoUrl} 
                alt="AGT Logo" 
                className={`w-10 h-10 object-contain opacity-90 ${logoTriedCount >= 3 ? 'hidden' : ''}`}
                onError={(e) => {
                  if (logoTriedCount === 0) {
                    setLogoTriedCount(1);
                    setLogoUrl('/AGTIcon.png');
                  } else if (logoTriedCount === 1) {
                    setLogoTriedCount(2);
                    setLogoUrl('/api/asset-proxy?id=1h9HvAGeru6Vo7PiWdLbXmGogD8TySnnz');
                  } else {
                    setLogoTriedCount(3);
                    const img = e.target as HTMLImageElement;
                    img.style.display = 'none';
                  }
                }}
              />
              {logoTriedCount >= 3 && (
                <div className="agt-fallback w-10 h-10 border border-agt-orange rounded-sm flex items-center justify-center shrink-0">
                  <span className="text-agt-orange font-bold text-[10px] tracking-tighter">AGT</span>
                </div>
              )}
            </div>
            <div className="flex flex-col">
              <h1 className="font-bold text-xs tracking-[0.2em] uppercase text-agt-orange">{t("Alliance of Galactic Travellers")}</h1>
              <span className="text-[9px] text-agt-orange uppercase tracking-[0.3em] font-bold">{t("AGT Region Analysis")}</span>
            </div>
          </div>
          
          <div className="flex items-center gap-6">
            <div className="hidden md:block text-[9px] text-agt-orange/30 tracking-widest font-mono">
              {t("STATUS: ")}<span className={
                loading ? 'text-yellow-500' :
                sheetUrl ? 'text-emerald-500' : 
                'text-red-500'
              }>
                {loading ? t('SYNCING') : sheetUrl ? t('CONNECTED') : t('DISCONNECTED')}
              </span>
            </div>
            <motion.button 
              onClick={() => setShowSettings(!showSettings)}
              className="p-2 hover:bg-[#FF0500]/5 rounded-lg transition-colors relative group text-[#FF0500]"
              title="Settings"
              id="settings-btn"
              whileHover={{ rotate: 180 }}
              transition={{ duration: 0.3 }}
              whileTap={{ scale: 0.9, rotate: -90 }}
            >
              <Settings className="w-5 h-5" />
              {!sheetUrl && (
                <span className="absolute top-0 right-0 w-1.5 h-1.5 bg-[#FF0500] rounded-full shadow-[0_0_5px_rgba(255,5,0,0.5)]"></span>
              )}
            </motion.button>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-16">
        <div className="flex flex-col gap-14">
          
          {/* Main Search Logic Container - centered aesthetic */}
          <div className="flex flex-col items-center space-y-10">
            <div className="w-full max-w-xl text-center space-y-6 flex flex-col items-center">
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <img 
                  src={starSystemsIcon} 
                  alt="Star Systems" 
                  className="w-12 h-12 object-contain" 
                  id="star-systems-icon-img"
                />
                <h2 className="text-4xl font-light tracking-tight text-[#FFB451]" id="main-title">
                  {t("AGT Region Analysis")}
                </h2>
              </div>
            </div>

            <div className="w-full max-w-2xl space-y-5 bg-[#111111] p-8 border-2 border-[#FF0500] rounded-2xl shadow-[0_0_20px_rgba(255,5,0,0.1)]">
              <div className="space-y-3">
                <div className="space-y-1 text-center select-none">
                  <p className="text-[#FFB451] text-[10px] font-bold tracking-widest uppercase">{t("Enter Region Name")}</p>
                </div>
                <div className="max-w-md mx-auto">
                  <AutocompleteInput
                    id="region-select"
                    value={selectedRegion}
                    onChange={(val) => {
                      setSelectedRegion(val);
                      if (data.length) {
                        findRecord(data, columns, val);
                      }
                    }}
                    suggestions={regionSuggestions}
                    placeholder={t("Enter/Choose Region Name...")}
                    icon={<Search className="h-5 w-5" />}
                  />
                </div>
              </div>

              <div className="flex justify-center pt-2">
                <button
                  onClick={handleSearch}
                  disabled={loading}
                  className="px-8 py-3 bg-[#E25530] text-white hover:bg-[#E25530]/90 disabled:opacity-50 transition-all font-mono uppercase tracking-widest text-xs font-bold rounded-lg border border-red-500 shadow-[0_0_15px_rgba(226,85,48,0.3)] hover:scale-[1.02]"
                >
                  {loading ? t("SCANNING DATABASE...") : t("INITIATE REGION SCAN")}
                </button>
              </div>
            </div>

            {error && (
              <div className="text-center p-3.5 bg-[#FF0500]/5 border border-[#FF0500] text-[#FFB451] rounded-full text-xs font-semibold tracking-wide w-full max-w-md">
                {error}
              </div>
            )}
          </div>

          <div className="space-y-12">
            
            {/* Settings Area - Beautiful pop-up window overlay */}
            <AnimatePresence>
              {showSettings && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm">
                  <div className="absolute inset-0" onClick={() => setShowSettings(false)} />
                  
                  <motion.div 
                    initial={{ scale: 0.95, opacity: 0 }}
                    animate={{ scale: 1, opacity: 1 }}
                    exit={{ scale: 0.95, opacity: 0 }}
                    className="w-full max-w-4xl max-h-[85vh] bg-[#111111] border-2 border-[#FF0500] rounded-2xl flex flex-col overflow-hidden relative shadow-[0_0_50px_rgba(255,5,0,0.25)]"
                    id="settings-popup-window"
                  >
                    {/* Header */}
                    <div className="p-6 border-b border-[#FF0500] bg-[#161616] flex items-center justify-between shrink-0">
                      <div className="flex items-center gap-2">
                        <Settings className="w-5 h-5 text-[#FF0500] animate-[spin_5s_linear_infinite]" />
                        <h2 className="text-md font-bold uppercase tracking-[0.2em] text-white">{t("System Settings Console")}</h2>
                      </div>
                      <button 
                        onClick={() => setShowSettings(false)}
                        className="px-5 py-2.5 bg-[#E25530] text-white border border-[#FF0500] rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-[#E25530]/90 transition"
                        id="settings-header-close"
                      >
                        {t("Close Settings")}
                      </button>
                    </div>

                    {/* Scrollable Content */}
                    <div className="p-8 overflow-y-auto flex-1 settings-scrollbar space-y-12 bg-[#0c0c0c]">
                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-12">

                        {/* Display Font Scale Section */}
                        <div className="space-y-4">
                          <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                            <Table className="w-3 h-3" />
                            {t("Display Font Scale")}
                          </h3>
                          <div className="space-y-2 relative">
                            <select
                              value={fontScale}
                              onChange={(e) => setFontScale(e.target.value)}
                              className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500] appearance-none"
                              id="fontscale-select"
                            >
                              {['1x', '1.5x', '2x', '2.5x', '3x'].map(scale => (
                                <option key={scale} value={scale}>
                                  {scale === '1x' ? t('1x (Default)') : scale}
                                </option>
                              ))}
                            </select>
                            <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-white/40">
                              <ChevronRight className="w-4 h-4 rotate-90" />
                            </div>
                          </div>
                        </div>

                        {/* Language Selection Section */}
                        <div className="space-y-4">
                          <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                            <Globe className="w-3 h-3" />
                            {t("Language Selection")}
                          </h3>
                          <div className="space-y-2 relative">
                            <select
                              value={language}
                              onChange={(e) => setLanguage(e.target.value as LanguageCode)}
                              className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500] appearance-none"
                              id="language-select"
                            >
                              {LANGUAGES.map(lang => (
                                <option key={lang.code} value={lang.code}>
                                  {lang.native}
                                </option>
                              ))}
                            </select>
                            <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-white/40">
                              <ChevronRight className="w-4 h-4 rotate-90" />
                            </div>
                          </div>
                        </div>

                        {/* AGT Anthem (Audio Section) */}
                        <div className="col-span-1 md:col-span-2 lg:col-span-3 pt-8 border-t border-[#FF0500]/20 space-y-4">
                          <div className="flex items-center justify-between">
                            <div className="space-y-1">
                              <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                                <Volume2 className="w-3 h-3" />
                                {t("AGT Anthem")}
                              </h3>
                            </div>
                             <button 
                              onClick={() => setAudioEnabled(!audioEnabled)}
                              className="flex items-center gap-3 px-6 py-3 rounded-xl border border-[#FF0500] bg-[#E25530] text-white transition-all text-[10px] uppercase tracking-widest font-bold hover:bg-[#E25530]/90"
                              id="audio-toggle-btn"
                            >
                              {audioEnabled ? <Volume2 className="w-3.5 h-3.5" /> : <VolumeX className="w-3.5 h-3.5" />}
                              {audioEnabled ? t('Active') : t('Muted')}
                            </button>
                          </div>
                        </div>

                        {/* System Database Section at the bottom of the scrollable content */}
                        <div className="col-span-1 md:col-span-2 lg:col-span-3 pt-8 border-t border-[#FF0500]/20 space-y-4">
                          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                            <div className="space-y-1">
                              <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                                <Database className="w-3 h-3" />
                                {t("System Database Settings")}
                              </h3>
                              <p className="text-[10px] text-[#FFB451]/60 font-mono tracking-wide italic">
                                {t("System database sync may take up to 5 minutes")}
                              </p>
                            </div>
                            <button 
                              onClick={() => {
                                fetchData();
                              }}
                              className="w-full sm:w-auto px-8 py-3.5 bg-[#E25530] border border-[#FF0500] text-white rounded-xl text-[10px] uppercase tracking-[0.1em] font-semibold hover:bg-[#E25530]/90 transition-colors shadow-[0_4px_20px_rgba(226,85,48,0.2)] active:scale-[0.98]"
                              id="resync-db-btn"
                            >
                              {t("Re-Sync System DB")}
                            </button>
                          </div>
                        </div>

                      </div>
                    </div>

                  </motion.div>
                </div>
              )}
            </AnimatePresence>

          {/* Results dashboard rendering area */}
          <div className="w-full">
            {error && (
              <div className="text-center p-4 bg-red-950/20 border border-red-400/30 text-red-400 font-mono text-xs rounded-xl uppercase tracking-wider max-w-xl mx-auto mb-6">
                {error}
              </div>
            )}

            <AnimatePresence mode="wait">
              {loading ? (
                <motion.div
                  key="loading"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="py-24 flex flex-col items-center justify-center text-center space-y-6"
                >
                  <div className="relative w-16 h-16 flex items-center justify-center">
                    <div className="absolute inset-0 rounded-full border-2 border-t-[#FFB451] border-[#FF0500]/20 animate-spin"></div>
                    <Database className="w-6 h-6 text-[#FFB451] animate-pulse" />
                  </div>
                  <div className="space-y-1">
                    <p className="text-sm font-bold uppercase tracking-[0.25em] text-[#FFB451]">{t("SCANNING AGT GALACTIC LEDGER...")}</p>
                    <p className="text-xs text-white/40 font-mono">{t("DECIPHERING ENCRYPTED SECTOR TELEMETRY")}</p>
                  </div>
                </motion.div>
              ) : matchedRecords.length > 0 && analysisStats ? (
                <motion.div
                  key="analysis_dashboard"
                  initial={{ opacity: 0, y: 40 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.98 }}
                  transition={{ duration: 0.5 }}
                  className="space-y-12"
                >
                  {/* Dashboard header card */}
                  <div className="p-8 bg-[#111111] border-2 border-[#FF0500] rounded-2xl flex flex-col md:flex-row md:items-center justify-between gap-6 shadow-[0_0_25px_rgba(255,5,0,0.06)]">
                    <div className="space-y-1.5">
                      <h3 className="text-2xl font-light text-white uppercase tracking-wider">
                        {t("REGION SCAN REPORT:")} <span className="font-bold text-[#FFB451]">{selectedRegion}</span>
                      </h3>
                      <p className="text-[10px] text-white/50 uppercase tracking-[0.25em] flex flex-wrap items-center gap-x-3 gap-y-1">
                        <span>{t("GALAXY:")} <strong className="text-[#FFB451]">{analysisStats.galaxyName}</strong></span>
                        <span className="text-white/20">|</span>
                        <span>{t("STATUS:")} <strong className="text-emerald-500 font-mono">{t("SECURE")}</strong></span>
                        <span className="text-white/20">|</span>
                        <span>{t("System Records:")} <strong className="text-[#FFB451] font-mono">{matchedRecords.length}</strong></span>
                      </p>
                    </div>
                    <div className="shrink-0 flex items-center gap-3">
                      <div className="px-4 py-2 border border-emerald-500/30 bg-emerald-950/20 rounded-lg flex items-center gap-2">
                        <div className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></div>
                        <span className="text-[9px] uppercase tracking-widest font-mono text-emerald-400">{t("SCAN COMPLETE")}</span>
                      </div>
                    </div>
                  </div>

                  {/* Section 1: Region Characteristics */}
                  <section className="space-y-6" id="section-1-characteristics">
                    <div className="border-b border-[#FF0500]/20 pb-2">
                      <h3 className="text-lg font-bold uppercase tracking-[0.2em] text-[#FFB451] flex items-center gap-3">
                        <span className="w-2.5 h-2.5 bg-[#FF0500]"></span>
                        {t("Section 1: Region Characteristics")}
                      </h3>
                    </div>

                    <div className="flex flex-col gap-6">
                      {/* Basic Facts Subsection */}
                      <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-3">
                        <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451]/80 border-b border-[#FF0500]/10 pb-1.5 flex items-center justify-between">
                          <span>{t("Basic Facts")}</span>
                          <span className="text-[10px] uppercase text-white/30 font-mono tracking-tighter">SUB_REF_FACTS</span>
                        </h4>
                        
                        <div className="space-y-1.5 font-mono text-sm">
                          <div className="flex justify-between items-center py-0.5 border-b border-[#FF0500]/5">
                            <span className="text-white/50 uppercase tracking-widest text-xs">{t("Region Name:")}</span>
                            <span className="font-bold text-[#FFB451] select-all">{selectedRegion}</span>
                          </div>
                          <div className="flex justify-between items-center py-0.5 border-b border-[#FF0500]/5">
                            <span className="text-white/50 uppercase tracking-widest text-xs">{t("Region Age:")}</span>
                            <span className="font-bold text-[#FFB451]">{analysisStats.regionAge}</span>
                          </div>
                          <div className="flex justify-between items-center py-0.5 border-b border-[#FF0500]/5">
                            <span className="text-white/50 uppercase tracking-widest text-xs">{t("Region Claim Status:")}</span>
                            <span className="font-bold text-white tracking-wide">{analysisStats.regionClaimStatus}</span>
                          </div>
                          <div className="flex justify-between items-center py-0.5 border-b border-[#FF0500]/5">
                            <span className="text-white/50 uppercase tracking-widest text-xs">{t("Estimated Normal Systems:")}</span>
                            <span className="font-bold text-[#FFB451]">{analysisStats.estimatedNormalSystems}</span>
                          </div>
                          <div className="flex justify-between items-center py-0.5 border-b border-[#FF0500]/5">
                            <span className="text-white/50 uppercase tracking-widest text-xs">{t("Galactic Coordinate Base Address:")}</span>
                            <span className="font-bold text-teal-400 select-all tracking-wider">{analysisStats.coordinateBaseAddress}</span>
                          </div>
                          <div className="flex justify-between items-center py-0.5">
                            <span className="text-white/50 uppercase tracking-widest text-xs">{t("Galaxy:")}</span>
                            <span className="font-bold text-[#FFB451]">{analysisStats.galaxyName}</span>
                          </div>
                        </div>
                      </div>

                      {/* Special Systems Subsection */}
                      <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-5">
                        <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451]/80 border-b border-[#FF0500]/10 pb-2 flex items-center justify-between">
                          <span>{t("Special Systems")}</span>
                          <span className="text-[10px] uppercase text-[#FFB451]/40 font-mono tracking-tighter">SUB_REF_ANOMALIES</span>
                        </h4>

                        <div className="grid grid-cols-2 gap-4 font-mono text-sm">
                          <div className="p-3.5 bg-black/50 border border-[#FF0500]/5 rounded-xl space-y-1 text-center">
                            <span className="text-xs uppercase text-white/40 tracking-wider block">{t("Black Hole")}</span>
                            <span className="font-bold text-red-500 block truncate text-sm" title={analysisStats.blackHoleSystem}>{analysisStats.blackHoleSystem}</span>
                          </div>
                          <div className="p-3.5 bg-black/50 border border-[#FF0500]/5 rounded-xl space-y-1 text-center">
                            <span className="text-xs uppercase text-white/40 tracking-wider block">{t("Atlas")}</span>
                            <span className="font-bold text-blue-500 block truncate text-sm" title={analysisStats.atlasSystem}>{analysisStats.atlasSystem}</span>
                          </div>
                          <div className="p-3.5 bg-black/50 border border-[#FF0500]/5 rounded-xl space-y-1 text-center">
                            <span className="text-xs uppercase text-white/40 tracking-wider block">{t("Shadow Star")}</span>
                            <span className="font-bold text-purple-400 block truncate text-sm" title={analysisStats.shadowSystem}>{analysisStats.shadowSystem}</span>
                          </div>
                          <div className="p-3.5 bg-black/50 border border-[#FF0500]/5 rounded-xl space-y-1 text-center">
                            <span className="text-xs uppercase text-white/40 tracking-wider block">{t("Reverse Shadow Star")}</span>
                            <span className="font-bold text-orange-400 block truncate text-sm" title={analysisStats.reverseShadowSystem}>{analysisStats.reverseShadowSystem}</span>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Factions Demographic Subsection */}
                    <div className="flex flex-col gap-6">
                      <div className="w-full font-sans">
                        <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-4 flex flex-col justify-between">
                          <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451] border-b border-[#FF0500]/10 pb-2">
                            {t("Demographics")}
                          </h4>
                          <div className="flex flex-col justify-center min-h-[300px]">
                            <SvgPieChart title="Factions" data={analysisStats.factionsPieData} hideLegend={true} radius={110} height={280} labelFontSize={12} hideSliceLabels={true} />
                          </div>
                          
                          {/* Faction Values Table shown in two columns below the pie chart */}
                          <div className="border-t border-[#FF0500]/10 pt-4 mt-2">
                            <span className="text-xs uppercase tracking-widest text-[#FFB451]/80 font-mono font-bold block mb-2">{t("Faction Breakdown")}</span>
                            <div className="grid grid-cols-2 gap-x-4 gap-y-1 text-xs font-mono">
                              {(() => {
                                const totalFactions = analysisStats.factionsPieData.reduce((acc, item) => acc + item.value, 0);
                                return analysisStats.factionsPieData.map((item, idx) => {
                                  const pct = totalFactions > 0 ? Math.round((item.value / totalFactions) * 100) : 0;
                                  return (
                                    <div key={idx} className="flex items-center justify-between py-1 border-b border-[#FF0500]/5">
                                      <div className="flex items-center gap-1.5 truncate">
                                        <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: item.color || '#FFB451' }}></span>
                                        <span className="font-semibold uppercase tracking-wide truncate" style={{ color: item.color || '#ffffff' }}>{item.label}</span>
                                      </div>
                                      <span className="font-bold shrink-0 ml-1" style={{ color: item.color || '#FFB451' }}>{item.value} ({pct}%)</span>
                                    </div>
                                  );
                                });
                              })()}
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Distributions Section */}
                      <div className="w-full p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-6 font-sans">
                        <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451] border-b border-[#FF0500]/10 pb-2">
                          {t("Distribution")}
                        </h4>
                        
                        <div className="flex flex-col gap-6 w-full">
                          {/* Star Color */}
                          <div className="flex flex-col gap-3 p-4 bg-black/20 border border-[#FF0500]/5 rounded-xl">
                            <div>
                              <SvgPieChart 
                                title="Star Color" 
                                data={analysisStats.starColorPieData} 
                                radius={105} 
                                height={260} 
                                labelFontSize={12} 
                                hideLegend={true} 
                                hideSliceLabels={true}
                              />
                            </div>
                            <div className="bg-black/30 border border-[#FF0500]/10 rounded-xl overflow-hidden mt-1">
                              <table className="w-full text-left font-mono text-xs border-collapse">
                                <thead>
                                  <tr className="border-b border-[#FF0500]/10 bg-black/40 text-neutral-400 select-none">
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Star Color Breakdown")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Systems")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Proportion")}</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-[#FF0500]/5">
                                  {analysisStats.starColorPieData.map((slice, i) => {
                                    const total = analysisStats.starColorPieData.reduce((acc, s) => acc + s.value, 0);
                                    const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                    return (
                                      <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                        <td className="py-2 px-4 flex items-center gap-2">
                                          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
                                          <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>{slice.label}</span>
                                        </td>
                                        <td className="py-2 px-4 text-[#FFB451] font-bold text-right text-sm">{slice.value}</td>
                                        <td className="py-2 px-4 text-white/50 text-right text-sm">{percent}%</td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* Dissonance */}
                          <div className="flex flex-col gap-3 p-4 bg-black/20 border border-[#FF0500]/5 rounded-xl">
                            <div>
                              <SvgPieChart 
                                title="Dissonance" 
                                data={analysisStats.dissonancePieData} 
                                radius={105} 
                                height={260} 
                                labelFontSize={12} 
                                hideLegend={true} 
                                hideSliceLabels={true}
                              />
                            </div>
                            <div className="bg-black/30 border border-[#FF0500]/10 rounded-xl overflow-hidden mt-1">
                              <table className="w-full text-left font-mono text-xs border-collapse">
                                <thead>
                                  <tr className="border-b border-[#FF0500]/10 bg-black/40 text-neutral-400 select-none">
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Dissonance Breakdown")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Systems")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Proportion")}</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-[#FF0500]/5">
                                  {analysisStats.dissonancePieData.map((slice, i) => {
                                    const total = analysisStats.dissonancePieData.reduce((acc, s) => acc + s.value, 0);
                                    const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                    return (
                                      <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                        <td className="py-2 px-4 flex items-center gap-2">
                                          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
                                          <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>{slice.label}</span>
                                        </td>
                                        <td className="py-2 px-4 text-[#FFB451] font-bold text-right text-sm">{slice.value}</td>
                                        <td className="py-2 px-4 text-white/50 text-right text-sm">{percent}%</td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* Giant Systems */}
                          <div className="flex flex-col gap-3 p-4 bg-black/20 border border-[#FF0500]/5 rounded-xl">
                            <div>
                              <SvgPieChart 
                                title="Giant Systems" 
                                data={analysisStats.giantPieData} 
                                radius={105} 
                                height={260} 
                                labelFontSize={12} 
                                hideLegend={true} 
                                hideSliceLabels={true}
                              />
                            </div>
                            <div className="bg-black/30 border border-[#FF0500]/10 rounded-xl overflow-hidden mt-1">
                              <table className="w-full text-left font-mono text-xs border-collapse">
                                <thead>
                                  <tr className="border-b border-[#FF0500]/10 bg-black/40 text-neutral-400 select-none">
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Giant Systems Breakdown")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Systems")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Proportion")}</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-[#FF0500]/5">
                                  {analysisStats.giantPieData.map((slice, i) => {
                                    const total = analysisStats.giantPieData.reduce((acc, s) => acc + s.value, 0);
                                    const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                    return (
                                      <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                        <td className="py-2 px-4 flex items-center gap-2">
                                          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
                                          <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>{slice.label}</span>
                                        </td>
                                        <td className="py-2 px-4 text-[#FFB451] font-bold text-right text-sm">{slice.value}</td>
                                        <td className="py-2 px-4 text-white/50 text-right text-sm">{percent}%</td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* Water Presence */}
                          <div className="flex flex-col gap-3 p-4 bg-black/20 border border-[#FF0500]/5 rounded-xl">
                            <div>
                              <SvgPieChart 
                                title="Water Presence" 
                                data={analysisStats.waterPieData} 
                                radius={105} 
                                height={260} 
                                labelFontSize={12} 
                                hideLegend={true} 
                                hideSliceLabels={true}
                              />
                            </div>
                            <div className="bg-black/30 border border-[#FF0500]/10 rounded-xl overflow-hidden mt-1">
                              <table className="w-full text-left font-mono text-xs border-collapse">
                                <thead>
                                  <tr className="border-b border-[#FF0500]/10 bg-black/40 text-neutral-400 select-none">
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Water Breakdown")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Systems")}</th>
                                    <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Proportion")}</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-[#FF0500]/5">
                                  {analysisStats.waterPieData.map((slice, i) => {
                                    const total = analysisStats.waterPieData.reduce((acc, s) => acc + s.value, 0);
                                    const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                    return (
                                      <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                        <td className="py-2 px-4 flex items-center gap-2">
                                          <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
                                          <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>{slice.label}</span>
                                        </td>
                                        <td className="py-2 px-4 text-[#FFB451] font-bold text-right text-sm">{slice.value}</td>
                                        <td className="py-2 px-4 text-white/50 text-right text-sm">{percent}%</td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Economics and Conflict Subsection */}
                    <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-6">
                      <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFD700] border-b border-[#FF0500]/10 pb-2">
                        {t("Economics and Conflict")}
                      </h4>

                      <div className="flex flex-col gap-6">
                        {/* Wealth Level Breakdown */}
                        <div className="p-6 bg-black/30 border border-[#FF0500]/10 rounded-2xl flex flex-col gap-4">
                          <div className="h-[320px]">
                            <SvgPieChart 
                              title="Wealth Level" 
                              data={analysisStats.wealthPieData} 
                              radius={105} 
                              height={260} 
                              labelFontSize={12} 
                              hideLegend={true} 
                              hideSliceLabels={true}
                            />
                          </div>
                          {/* Wealth Level Breakdown Table/Legend */}
                          <div className="bg-black/40 border border-[#FF0500]/10 rounded-xl overflow-hidden mt-1">
                            <table className="w-full text-left font-mono text-xs border-collapse">
                              <thead>
                                <tr className="border-b border-[#FF0500]/10 bg-black/30 text-neutral-400 select-none">
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Wealth Level Breakdown")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.wealthPieData.map((slice, i) => {
                                  const total = analysisStats.wealthPieData.reduce((acc, s) => acc + s.value, 0);
                                  const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                  return (
                                    <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                      <td className="py-2 px-4 flex items-center gap-2">
                                        <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
                                        <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>
                                          {slice.label} {slice.value} ({percent}%)
                                        </span>
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Economy Type Component */}
                        <div className="p-6 bg-black/30 border border-[#FF0500]/10 rounded-2xl flex flex-col gap-4">
                          <div className="h-[320px]">
                            <SvgPieChart 
                              title="Economy Type" 
                              data={analysisStats.economyTypePieData} 
                              radius={105} 
                              height={260} 
                              labelFontSize={12} 
                              hideLegend={true} 
                              hideSliceLabels={true}
                            />
                          </div>
                          {/* Separated legend in a three-column table with proportion percentages */}
                          <div className="bg-black/40 border border-[#FF0500]/10 rounded-xl overflow-hidden">
                            <table className="w-full text-left font-mono text-xs border-collapse">
                              <thead>
                                <tr className="border-b border-[#FF0500]/10 bg-black/30 text-neutral-400 select-none">
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Economy Type")}</th>
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Systems")}</th>
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Proportion")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.economyTypePieData.map((slice, i) => {
                                  const total = analysisStats.economyTypePieData.reduce((acc, s) => acc + s.value, 0);
                                  const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                  return (
                                    <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                      <td className="py-2 px-4 flex items-center gap-2">
                                        <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color || '#A855F7' }}></span>
                                        <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>{slice.label}</span>
                                      </td>
                                      <td className="py-2 px-4 text-[#FFB451] font-bold text-right text-sm">{slice.value}</td>
                                      <td className="py-2 px-4 text-white/50 text-right text-sm">{percent}%</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Conflict Level Component */}
                        <div className="p-6 bg-black/30 border border-[#FF0500]/10 rounded-2xl flex flex-col gap-4">
                          <div className="h-[320px]">
                            <SvgPieChart 
                              title="Conflict Level" 
                              data={analysisStats.conflictPieData} 
                              radius={105} 
                              height={260} 
                              labelFontSize={12} 
                              hideLegend={true} 
                              hideSliceLabels={true}
                            />
                          </div>
                          {/* Conflict Level Breakdown Table */}
                          <div className="bg-black/40 border border-[#FF0500]/10 rounded-xl overflow-hidden mt-1">
                            <table className="w-full text-left font-mono text-xs border-collapse">
                              <thead>
                                <tr className="border-b border-[#FF0500]/10 bg-black/30 text-neutral-400 select-none">
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest">{t("Conflict Level Breakdown")}</th>
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Systems")}</th>
                                  <th className="py-2 px-4 uppercase font-bold tracking-widest text-right">{t("Proportion")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.conflictPieData.map((slice, i) => {
                                  const total = analysisStats.conflictPieData.reduce((acc, s) => acc + s.value, 0);
                                  const percent = total > 0 ? Math.round((slice.value / total) * 100) : 0;
                                  return (
                                    <tr key={i} className="hover:bg-[#E25530]/5 transition-colors">
                                      <td className="py-2 px-4 flex items-center gap-2">
                                        <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slice.color }}></span>
                                        <span className="font-semibold uppercase tracking-wider text-sm" style={{ color: slice.color }}>{slice.label}</span>
                                      </td>
                                      <td className="py-2 px-4 text-[#FFB451] font-bold text-right text-sm">{slice.value}</td>
                                      <td className="py-2 px-4 text-white/50 text-right text-sm">{percent}%</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>

                      {/* Buy and sell text boxes / font size increased */}
                      <div className="flex flex-col gap-4 font-mono text-sm">
                        <div className="p-5 bg-black/50 border border-[#FF0500]/10 rounded-xl space-y-4">
                          <div className="space-y-1 border-b border-[#FF0500]/5 pb-2">
                            <span className="text-xs text-white/40 uppercase tracking-widest block">{t("Highest Buy System:")} {analysisStats.highestBuyVal ? `[${analysisStats.highestBuyVal}]` : ''}</span>
                            <p className="text-[#FFB451] font-bold text-base tracking-wide leading-relaxed">{analysisStats.highestBuy}</p>
                          </div>
                          <div className="space-y-1">
                            <span className="text-xs text-white/40 uppercase tracking-widest block">{t("Lowest Buy System:")} {analysisStats.lowestBuyVal ? `[${analysisStats.lowestBuyVal}]` : ''}</span>
                            <p className="text-[#FFB451] font-bold text-base tracking-wide leading-relaxed">{analysisStats.lowestBuy}</p>
                          </div>
                        </div>

                        <div className="p-5 bg-black/50 border border-[#FF0500]/10 rounded-xl space-y-4">
                          <div className="space-y-1 border-b border-[#FF0500]/5 pb-2">
                            <span className="text-xs text-white/40 uppercase tracking-widest block">{t("Highest Sell System:")} {analysisStats.highestSellVal ? `[${analysisStats.highestSellVal}]` : ''}</span>
                            <p className="text-[#FFB451] font-bold text-base tracking-wide leading-relaxed">{analysisStats.highestSell}</p>
                          </div>
                          <div className="space-y-1">
                            <span className="text-xs text-white/40 uppercase tracking-widest block">{t("Lowest Sell System:")} {analysisStats.lowestSellVal ? `[${analysisStats.lowestSellVal}]` : ''}</span>
                            <p className="text-[#FFB451] font-bold text-base tracking-wide leading-relaxed">{analysisStats.lowestSell}</p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </section>

                  {/* Section 2: System Claiming */}
                  <section className="space-y-6" id="section-2-claiming">
                    <div className="border-b border-[#FF0500]/20 pb-2">
                      <h3 className="text-lg font-bold uppercase tracking-[0.2em] text-[#FFB451] flex items-center gap-3">
                        <span className="w-2.5 h-2.5 bg-[#FF0500]"></span>
                        {t("Section 2: System Claiming")}
                      </h3>
                    </div>

                    {/* Civ/Company Claims Subsection / Stacked Vertically, No Multi-Column Layout */}
                    <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-6">
                      <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451] border-b border-[#FF0500]/10 pb-2">
                        {t("Civ/Company Claims")}
                      </h4>

                      <div className="flex flex-col gap-6 w-full items-stretch animate-none">
                        <div className="h-[300px] w-full flex items-center justify-center bg-black/20 rounded-xl border border-[#FF0500]/5 p-2">
                          <SvgPieChart 
                            title="System Claims" 
                            data={analysisStats.systemClaimsPieData} 
                            useSpecialColors={true} 
                            hideLegend={true}
                            radius={92} 
                            height={256} 
                            labelFontSize={12} 
                            hideSliceLabels={true}
                          />
                        </div>

                        <div className="bg-black/50 border border-[#FF0500]/10 rounded-2xl overflow-hidden shadow-inner flex flex-col justify-between w-full">
                          <div className="p-4 border-b border-[#FF0500]/10 bg-[#161616]">
                            <span className="text-xs font-mono tracking-widest uppercase text-white">{t("Civilization Claims")}</span>
                          </div>
                          
                          <div className="overflow-y-auto max-h-[350px] settings-scrollbar">
                            <table className="w-full text-left font-mono text-sm border-collapse">
                              <thead>
                                <tr className="border-b border-[#FF0500]/15 bg-black/40 text-neutral-400 select-none">
                                  <th className="py-1.5 px-4 font-bold tracking-widest uppercase text-xs">{t("Civilization Label")}</th>
                                  <th className="py-1.5 px-4 font-bold tracking-widest uppercase text-xs text-right">{t("Systems")}</th>
                                  <th className="py-1.5 px-4 font-bold tracking-widest uppercase text-xs text-right">{t("Proportion")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.claimsTableData.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-[#E25530]/5 transition-colors">
                                    <td className="py-1 px-4 font-semibold truncate" style={{ color: row.color }}>{row.civ}</td>
                                    <td className="py-1 px-4 text-[#FFB451] font-bold text-right">{row.count}</td>
                                    <td className="py-1 px-4 text-white/50 text-right">{row.percent}%</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    </div>
                  </section>

                  {/* Section 3: Discovery and Surveyors */}
                  <section className="space-y-8" id="section-3-discovery">
                    <div className="border-b border-[#FF0500]/20 pb-2">
                      <h3 className="text-lg font-bold uppercase tracking-[0.2em] text-[#FFB451] flex items-center gap-3">
                        <span className="w-2.5 h-2.5 bg-[#FF0500]"></span>
                        {t("Section 3: Discovery and Surveyors")}
                      </h3>
                    </div>

                    {/* Discoverers Subsection / Stacked Vertically, No Multi-Column Layout */}
                    <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-6">
                      <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451] border-b border-[#FF0500]/10 pb-2 flex items-center justify-between">
                        <span>{t("Discoverers")}</span>
                      </h4>

                      <div className="flex flex-col gap-6 font-sans">
                        <div className="space-y-4 flex flex-col items-center w-full">
                          <h5 className="text-sm font-bold uppercase tracking-widest text-white/80 text-center w-full">{t("Top 10 Discoverers")}</h5>
                          <div className="bg-black/50 border border-[#FF0500]/10 rounded-xl overflow-hidden font-mono text-sm w-full">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="border-b border-[#FF0500]/15 bg-black/30 text-neutral-400 select-none">
                                  <th className="py-1 px-4 uppercase font-bold text-xs tracking-widest">{t("Discoverer Name")}</th>
                                  <th className="py-1 px-4 uppercase font-bold text-xs tracking-widest text-right">{t("Systems")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5 text-sm">
                                {analysisStats.top10Discoverers.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-[#E25530]/5">
                                    <td className="py-1 px-4 text-white truncate">{row.name}</td>
                                    <td className="py-1 px-4 text-[#FFB451] font-bold text-right">{row.count}</td>
                                  </tr>
                                ))}
                                {analysisStats.top10Discoverers.length === 0 && (
                                  <tr>
                                    <td colSpan={2} className="py-4 text-center text-white/20 italic">{t("No Data Found")}</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        <div className="h-[460px] min-h-[460px] flex flex-col justify-between w-full bg-black/20 rounded-xl border border-[#FF0500]/5 p-4 items-center">
                          <h5 className="text-sm font-bold uppercase tracking-widest text-white/80 text-center w-full pb-2">{t("Discoverer Distribution")}</h5>
                          <div className="flex-1 w-full max-w-lg mx-auto flex items-center justify-center">
                            <SvgPieChart title="Discoverers Breakdown" data={analysisStats.discoverersPieData} hideLegend={true} radius={85} height={240} labelFontSize={12} />
                          </div>
                        </div>
                      </div>

                      <div className="flex flex-col gap-6 border-t border-[#FF0500]/5 pt-6">
                        <div className="space-y-3">
                          <h5 className="text-xs font-bold uppercase tracking-widest text-emerald-400 font-mono">{t("10 Earliest Discovery Dates")}</h5>
                          <div className="bg-black/50 border border-[#FF0500]/10 rounded-xl overflow-hidden font-mono text-xs">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="bg-black/30 text-neutral-400 border-b border-[#FF0500]/15 text-xs">
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("System")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Date")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Discoverer")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.earliestDiscoveryDates.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-emerald-950/10">
                                    <td className="py-2 px-3 text-[#FFB451] font-semibold truncate max-w-[120px]">{row.system}</td>
                                    <td className="py-2 px-3 text-white/60">{row.dateStr}</td>
                                    <td className="py-2 px-3 text-white truncate max-w-[120px]">{row.name}</td>
                                  </tr>
                                ))}
                                {analysisStats.earliestDiscoveryDates.length === 0 && (
                                  <tr>
                                    <td colSpan={3} className="py-4 text-center text-white/20 italic">{t("No Entry")}</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        <div className="space-y-3">
                          <h5 className="text-xs font-bold uppercase tracking-widest text-[#FF0500] font-mono">{t("10 Most Recent Discovery Dates")}</h5>
                          <div className="bg-black/50 border border-[#FF0500]/10 rounded-xl overflow-hidden font-mono text-xs">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="bg-black/30 text-neutral-400 border-b border-[#FF0500]/15 text-xs">
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("System")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Date")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Discoverer")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.mostRecentDiscoveryDates.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-red-950/10">
                                    <td className="py-2 px-3 text-[#FFB451] font-semibold truncate max-w-[120px]">{row.system}</td>
                                    <td className="py-2 px-3 text-white/60">{row.dateStr}</td>
                                    <td className="py-2 px-3 text-white truncate max-w-[120px]">{row.name}</td>
                                  </tr>
                                ))}
                                {analysisStats.mostRecentDiscoveryDates.length === 0 && (
                                  <tr>
                                    <td colSpan={3} className="py-4 text-center text-white/20 italic">{t("No Entry")}</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Surveyor Reference Subsection / Stacked Vertically, No Multi-Column Layout */}
                    <div className="p-6 bg-black/40 border border-[#FF0500]/10 rounded-2xl space-y-6">
                      <h4 className="text-sm font-bold uppercase tracking-widest text-[#FFB451] border-b border-[#FF0500]/10 pb-2 flex items-center justify-between">
                        <span>{t("Surveyors")}</span>
                      </h4>

                      <div className="flex flex-col gap-6 font-sans">
                        <div className="space-y-4 flex flex-col items-center w-full">
                          <h5 className="text-sm font-bold uppercase tracking-widest text-[#FF0500] text-center w-full">{t("Top 10 Surveyors")}</h5>
                          <div className="bg-black/50 border border-[#FF0500]/10 rounded-xl overflow-hidden font-mono text-sm w-full">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="border-b border-[#FF0500]/15 bg-black/30 text-neutral-400 select-none">
                                  <th className="py-1 px-4 uppercase font-bold text-xs tracking-widest">{t("Surveyor Name")}</th>
                                  <th className="py-1 px-4 uppercase font-bold text-xs tracking-widest text-right">{t("Systems")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5 text-sm">
                                {analysisStats.top10Surveyors.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-[#E25530]/5">
                                    <td className="py-1 px-4 text-white truncate">{row.name}</td>
                                    <td className="py-1 px-4 text-[#FFB451] font-bold text-right">{row.count}</td>
                                  </tr>
                                ))}
                                {analysisStats.top10Surveyors.length === 0 && (
                                  <tr>
                                    <td colSpan={2} className="py-4 text-center text-white/20 italic">{t("No Data Found")}</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        <div className="h-[460px] min-h-[460px] flex flex-col justify-between w-full bg-black/20 rounded-xl border border-[#FF0500]/5 p-4 items-center">
                          <h5 className="text-sm font-bold uppercase tracking-widest text-white/80 text-center w-full pb-2">{t("Surveyor Distribution")}</h5>
                          <div className="flex-1 w-full max-w-lg mx-auto flex items-center justify-center">
                            <SvgPieChart title="Surveyors Breakdown" data={analysisStats.surveyorsPieData} hideLegend={true} radius={85} height={240} labelFontSize={12} />
                          </div>
                        </div>
                      </div>

                      <div className="flex flex-col gap-6 border-t border-[#FF0500]/5 pt-6">
                        <div className="space-y-3">
                          <h5 className="text-xs font-bold uppercase tracking-widest text-emerald-400 font-mono">{t("10 Oldest Survey Dates")}</h5>
                          <div className="bg-black/50 border border-[#FF0500]/10 rounded-xl overflow-hidden font-mono text-xs">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="bg-black/30 text-neutral-400 border-b border-[#FF0500]/15 text-xs">
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("System")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Date")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Surveyor")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.earliestSurveyorDates.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-emerald-950/10">
                                    <td className="py-2 px-3 text-[#FFB451] font-semibold truncate max-w-[120px]">{row.system}</td>
                                    <td className="py-2 px-3 text-white/60">{row.dateStr}</td>
                                    <td className="py-2 px-3 text-white truncate max-w-[120px]">{row.name}</td>
                                  </tr>
                                ))}
                                {analysisStats.earliestSurveyorDates.length === 0 && (
                                  <tr>
                                    <td colSpan={3} className="py-4 text-center text-white/20 italic">{t("No Entry")}</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        <div className="space-y-3">
                          <h5 className="text-xs font-bold uppercase tracking-widest text-[#FF0500] font-mono">{t("10 Most Recent Surveyor Dates")}</h5>
                          <div className="bg-black/50 border border-[#FF0500]/10 rounded-xl overflow-hidden font-mono text-xs">
                            <table className="w-full text-left">
                              <thead>
                                <tr className="bg-black/30 text-neutral-400 border-b border-[#FF0500]/15 text-xs">
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("System")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Date")}</th>
                                  <th className="py-2 px-3 text-[10px] tracking-wider uppercase font-bold">{t("Surveyor")}</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-[#FF0500]/5">
                                {analysisStats.mostRecentSurveyorDates.map((row, idx) => (
                                  <tr key={idx} className="hover:bg-red-950/10">
                                    <td className="py-2 px-3 text-[#FFB451] font-semibold truncate max-w-[120px]">{row.system}</td>
                                    <td className="py-2 px-3 text-white/60">{row.dateStr}</td>
                                    <td className="py-2 px-3 text-white truncate max-w-[120px]">{row.name}</td>
                                  </tr>
                                ))}
                                {analysisStats.mostRecentSurveyorDates.length === 0 && (
                                  <tr>
                                    <td colSpan={3} className="py-4 text-center text-white/20 italic">{t("No Entry")}</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    </div>
                  </section>

                  {/* Section 4: Additional Information */}
                  <section className="space-y-6" id="section-4-additional">
                    <div className="border-b border-[#FF0500]/20 pb-2">
                      <h3 className="text-lg font-bold uppercase tracking-[0.2em] text-[#FFB451] flex items-center gap-3">
                        <span className="w-2.5 h-2.5 bg-[#FF0500]"></span>
                        {t("Section 4: Additional Information")}
                      </h3>
                    </div>

                    <div className="p-8 bg-black/40 border border-[#FF0500]/10 rounded-2xl">
                      <p className="text-white/60 text-xs font-mono uppercase tracking-widest text-center">{t("For Future use")}</p>
                    </div>
                  </section>

                  {/* PDF Generation Button Segment */}
                  <div className="pt-6 flex justify-center pb-8">
                    <button
                      onClick={downloadRegionAnalysisReportPdf}
                      style={{ backgroundColor: '#E25530' }}
                      className="px-8 py-3.5 text-white font-sans font-bold uppercase tracking-widest text-xs rounded-full shadow-[0_0_15px_rgba(226,85,48,0.3)] hover:brightness-110 active:scale-95 transition-all flex items-center gap-2 cursor-pointer border border-[#FF0500]/10"
                      id="create-pdf-report-btn"
                    >
                      <FileText className="w-4 h-4 text-white" />
                      {t("Create PDF Report")}
                    </button>
                  </div>

                </motion.div>
              ) : !loading && (
                <motion.div 
                  key="empty"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="py-32 flex flex-col items-center justify-center text-center space-y-6 opacity-40 border-2 border-dashed border-[#FF0500]/25 rounded-2xl bg-black/10"
                >
                  <div className="w-16 h-16 rounded-full border border-[#FFB451]/20 flex items-center justify-center">
                    <Database className="w-6 h-6 text-[#FFB451]" />
                  </div>
                  <div className="space-y-1.5 max-w-sm">
                    <p className="text-sm font-bold uppercase tracking-[0.2em] text-[#FFB451]">{t("CONSOLE INITIAL DECRYPT READY")}</p>
                    <p className="text-xs text-white/50 leading-relaxed">{t("Please select an active region name to execute automated telemetry scan procedures.")}</p>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>
      </div>
    </main>

      {/* Footer Area */}
      <footer className="bg-[#FFB451] mt-auto">
        <div className="max-w-5xl mx-auto px-6 py-12 flex flex-col items-center gap-6 text-black">
          <div className="flex flex-wrap justify-center items-center gap-y-2 text-[10px] uppercase tracking-[0.2em] font-bold">
            <a href="https://www.nms-agt.com" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Home</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/about-the-agt" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">About</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/team" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Team</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/contribute" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Contribute</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/agt-galactic-archives" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Galactic Archives</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/engage" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Engage</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/agt-navi" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">AGT NAVI</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/terms" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Terms</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/support" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Support</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/terms/copyright" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Copyright</a>
          </div>
          <p className="text-[9px] font-mono uppercase tracking-[0.3em] font-bold">&copy; 2026 {t("Alliance of Galactic Travellers")}</p>
        </div>
      </footer>

      {/* Background Audio */}
      <audio 
        ref={audioRef}
        src={agtAnthem}
        loop
        preload="auto"
      />

      {/* PDF Error Modal Pop-up */}
      <AnimatePresence>
        {pdfErrorMsg && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-black/90 backdrop-blur-sm">
            <div className="absolute inset-0" onClick={() => setPdfErrorMsg(null)} />
            
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="w-full max-w-lg bg-[#111111] border-2 border-[#FF0500] rounded-2xl p-8 flex flex-col relative shadow-[0_0_50px_rgba(255,5,0,0.4)] text-center space-y-6"
              id="pdf-error-popup"
            >
              <div className="mx-auto w-16 h-16 rounded-full bg-[#FF0500]/10 border-2 border-[#FF0500] flex items-center justify-center animate-pulse">
                <AlertCircle className="w-8 h-8 text-[#FF0500]" />
              </div>
              
              <div className="space-y-2">
                <h3 className="text-xl font-bold uppercase tracking-widest text-[#FF0500]" id="pdf-error-title">
                  Report Exceeds Limit
                </h3>
                <p className="text-xs text-[#FFB451] font-mono leading-relaxed animate-pulse" id="pdf-error-message">
                  {pdfErrorMsg}
                </p>
              </div>

              <div className="pt-2">
                <button 
                  onClick={() => setPdfErrorMsg(null)}
                  className="px-10 py-3.5 bg-[#E25530] border-2 border-[#FF0500] text-white rounded-full font-black text-[10px] uppercase tracking-widest hover:bg-[#E25530]/90 active:scale-[0.96] transition-all shadow-[0_4px_15px_rgba(255,5,0,0.2)]"
                  id="pdf-error-close-btn"
                >
                  Dismiss
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Search Error Modal Pop-up */}
      <AnimatePresence>
        {popupError && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-black/90 backdrop-blur-sm">
            <div className="absolute inset-0" onClick={() => setPopupError(null)} />
            
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="w-full max-w-lg bg-[#111111] border-2 border-[#FF0500] rounded-2xl p-8 flex flex-col relative shadow-[0_0_50px_rgba(255,5,0,0.4)] text-center space-y-6"
              id="search-error-popup"
            >
              <div className="mx-auto w-16 h-16 rounded-full bg-[#FF0500]/10 border-2 border-[#FF0500] flex items-center justify-center theme-glow">
                <AlertCircle className="w-8 h-8 text-[#FF0500] animate-bounce" />
              </div>
              
              <div className="space-y-2">
                <h3 className="text-xl font-bold uppercase tracking-widest text-[#FF0500]" id="search-error-title">
                  {t("User must enter a valid region name")}
                </h3>
              </div>

              <div className="pt-2">
                <button 
                  onClick={() => setPopupError(null)}
                  className="px-10 py-3.5 bg-[#E25530] border-2 border-[#FF0500] text-white rounded-full font-black text-[10px] uppercase tracking-widest hover:bg-[#E25530]/90 active:scale-[0.96] transition-all shadow-[0_4px_15px_rgba(255,5,0,0.2)]"
                  id="search-error-close-btn"
                >
                  Dismiss
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

