import React, {
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import * as XLSX from 'xlsx';

const DEFAULT_USERS = [
  '杨志军',
  '尹涛',
  '罗杰耀',
  '刘章凯',
  '郭永',
  '邓永明',
  '李建力',
  '拉平',
  '加布',
  '李保',
  '康晓波',
  '赵宏飞',
  '孙师傅',
  '普次师傅',
  '杨师傅',
].map((name, index) => ({
  id: index + 1,
  name,
}));

interface Row {
  id: number;
  name: string;
  baozi: number | '';
  mantou: number | '';
  porridge: number | '';
  egg: number | '';
  updatedAt: string;
}

const BREAKFAST_FIELDS = [
  { key: 'baozi', label: '包子', unit: '个' },
  { key: 'mantou', label: '馒头', unit: '个' },
  { key: 'porridge', label: '稀饭', unit: '碗' },
  { key: 'egg', label: '鸡蛋', unit: '个' },
] as const;

type FieldKey = (typeof BREAKFAST_FIELDS)[number]['key'];

const createEmptyRow = (user: { id: number; name: string }): Row => ({
  ...user,
  baozi: '',
  mantou: '',
  porridge: '',
  egg: '',
  updatedAt: '',
});

const createDefaultData = () =>
  DEFAULT_USERS.map((user) => createEmptyRow(user));

const getDateKey = () => {
  const now = new Date();
  return [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getDate()).padStart(2, '0'),
  ].join('-');
};

const APP_VERSION = 'v2';

const getStorageKey = () => `breakfast-${getDateKey()}-${APP_VERSION}`;
const getResetKey = () => `breakfast-reset-${getDateKey()}-${APP_VERSION}`;

const safeParse = <T,>(value: string, fallback: T): T => {
  try {
    return JSON.parse(value);
  } catch {
    return fallback;
  }
};

const loadRows = () => {
  if (typeof window === 'undefined') {
    return createDefaultData();
  }
  const cache = localStorage.getItem(getStorageKey());
  return cache ? safeParse(cache, createDefaultData()) : createDefaultData();
};

const loadHistory = () => {
  if (typeof window === 'undefined') {
    return [];
  }
  const cache = localStorage.getItem('breakfast-history');
  return cache ? safeParse(cache, []) : [];
};

export default function App() {
  const [rows, setRows] = useState<Row[]>(loadRows);
  const [history, setHistory] = useState<{ id: number; date: string; rows: Row[]; totals: Record<FieldKey, number> }[]>(loadHistory);
  const [search, setSearch] = useState('');
  const [lastArchive, setLastArchive] = useState('');

  const rowsRef = useRef(rows);
  const historyRef = useRef(history);

  useEffect(() => {
    rowsRef.current = rows;
  }, [rows]);

  useEffect(() => {
    historyRef.current = history;
  }, [history]);

  useEffect(() => {
    localStorage.setItem(getStorageKey(), JSON.stringify(rows));
  }, [rows]);

  const totals = useMemo(() => {
    return BREAKFAST_FIELDS.reduce(
      (acc, field) => {
        acc[field.key] = rows.reduce(
          (sum: number, row: Row) => sum + (Number(row[field.key]) || 0),
          0,
        );
        return acc;
      },
      {} as Record<FieldKey, number>,
    );
  }, [rows]);

  const filledCount = useMemo(() => {
    return rows.filter((row: Row) =>
      BREAKFAST_FIELDS.some((field) => row[field.key] !== ''),
    ).length;
  }, [rows]);

  const filteredRows = useMemo(() => {
    return rows.filter((row: Row) =>
      row.name.toLowerCase().includes(search.toLowerCase()),
    );
  }, [rows, search]);

  useEffect(() => {
    const autoArchive = () => {
      const now = new Date();
      const currentHour = now.getHours();
      const resetKey = getResetKey();

      if (currentHour >= 12 && !window.localStorage.getItem(resetKey)) {
        const currentRows = rowsRef.current;
        const archiveData = {
          id: Date.now(),
          date: now.toLocaleString(),
          rows: currentRows,
          totals,
        };
        const updatedHistory = [archiveData, ...historyRef.current];
        localStorage.setItem('breakfast-history', JSON.stringify(updatedHistory));
        localStorage.setItem(resetKey, 'done');
        const emptyRows = createDefaultData();
        localStorage.setItem(getStorageKey(), JSON.stringify(emptyRows));
        setHistory(updatedHistory);
        setRows(emptyRows);
        setLastArchive(now.toLocaleString());
      }
    };

    autoArchive();
    const timer = window.setInterval(autoArchive, 30000);
    return () => clearInterval(timer);
  }, [totals]);

  const updateField = (id: number, key: FieldKey, value: string) => {
    const sanitized = value === '' ? '' : Math.max(0, Number(value));
    setRows((prev: Row[]) =>
      prev.map((row: Row) =>
        row.id === id
          ? { ...row, [key]: sanitized, updatedAt: new Date().toLocaleTimeString() }
          : row,
      ),
    );
  };

  const clearToday = () => {
    setRows(createDefaultData());
  };

  const exportExcel = () => {
    const exportRows = rows.map((row: Row) => ({
      '姓名': row.name,
      '包子': row.baozi || 0,
      '馒头': row.mantou || 0,
      '稀饭': row.porridge || 0,
      '鸡蛋': row.egg || 0,
      '更新时间': row.updatedAt || '-',
    }));

    exportRows.push({
      '姓名': '总计',
      '包子': totals.baozi,
      '馒头': totals.mantou,
      '稀饭': totals.porridge,
      '鸡蛋': totals.egg,
      '更新时间': '-',
    });

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, '早餐统计');
    XLSX.writeFile(workbook, `早餐统计-${getDateKey()}.xlsx`);
  };

  const copyWechatText = async () => {
    const filled = rows.filter((row: Row) =>
      BREAKFAST_FIELDS.some((field) => row[field.key] !== ''),
    );
    const content = filled
      .map(
        (row: Row) =>
          `${row.name}｜包子:${row.baozi || 0} 馒头:${row.mantou || 0} 稀饭:${row.porridge || 0} 鸡蛋:${row.egg || 0}`,
      )
      .join('\n');

    const result = `早餐统计\n\n${content}\n\n总计：\n包子 ${totals.baozi}\n馒头 ${totals.mantou}\n稀饭 ${totals.porridge}\n鸡蛋 ${totals.egg}`;

    await navigator.clipboard.writeText(result);
    alert('已复制，可直接发送微信群');
  };

  return (
    <div className="min-h-screen bg-stone-50 p-3 md:p-6">
      <div className="mx-auto max-w-3xl rounded-2xl md:rounded-[32px] bg-white shadow-2xl overflow-hidden border border-slate-200">
        {/* Header */}
        <div className="bg-slate-900 text-white p-5 md:p-8">
          <h1 className="text-2xl md:text-4xl font-bold tracking-tight">
            早餐统计
          </h1>

          <div className="grid grid-cols-4 gap-2 md:gap-4 mt-5">
            {BREAKFAST_FIELDS.map((field) => (
              <div
                key={field.key}
                className="bg-white/10 rounded-xl md:rounded-2xl px-3 py-3 md:px-5 md:py-4 text-center"
              >
                <div className="text-xs md:text-sm text-slate-300">
                  {field.label}
                </div>
                <div className="text-xl md:text-3xl font-bold mt-1 md:mt-2">
                  {totals[field.key]}
                </div>
                <div className="text-[10px] md:text-xs text-slate-400">
                  {field.unit}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Toolbar */}
        <div className="p-3 md:p-6 border-b border-slate-100 bg-slate-50 space-y-3">
          <div className="flex flex-wrap gap-2 md:gap-3">
            <button
              onClick={copyWechatText}
              className="rounded-xl md:rounded-2xl bg-slate-900 text-white px-4 py-2.5 md:px-6 md:py-3 text-sm md:text-base font-medium active:scale-95 transition-all"
            >
              复制微信群统计
            </button>
            <button
              onClick={exportExcel}
              className="rounded-xl md:rounded-2xl bg-emerald-600 text-white px-4 py-2.5 md:px-6 md:py-3 text-sm md:text-base font-medium active:scale-95 transition-all"
            >
              导出 Excel
            </button>
          </div>

          <div className="flex gap-3 items-center">
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="搜索姓名..."
              className="flex-1 rounded-xl border border-slate-200 px-4 py-2.5 text-sm outline-none focus:ring-2 focus:ring-slate-400 bg-white"
            />
            <div className="text-sm text-slate-500 whitespace-nowrap">
              已填 <span className="font-bold text-slate-800">{filledCount}</span>/{rows.length}
            </div>
          </div>
        </div>

        {/* Card list */}
        <div className="p-3 md:p-6 space-y-2 md:space-y-3 max-h-[60vh] overflow-y-auto">
          {filteredRows.map((row) => {
            const isFilled = BREAKFAST_FIELDS.some(
              (field) => row[field.key as FieldKey] !== '',
            );
            return (
              <div
                key={row.id}
                className={`rounded-xl md:rounded-2xl border p-3 md:p-4 transition-all ${
                  isFilled
                    ? 'border-emerald-200 bg-emerald-50/50'
                    : 'border-slate-200 bg-white'
                }`}
              >
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-slate-800 text-sm md:text-base">
                      {row.name}
                    </span>
                    {isFilled ? (
                      <span className="text-[10px] md:text-xs rounded-full bg-emerald-500 text-white px-2 py-0.5 font-medium">
                        已填
                      </span>
                    ) : (
                      <span className="text-[10px] md:text-xs rounded-full bg-slate-300 text-slate-600 px-2 py-0.5 font-medium">
                        未填
                      </span>
                    )}
                  </div>
                  {row.updatedAt && (
                    <span className="text-[10px] md:text-xs text-slate-400">
                      {row.updatedAt}
                    </span>
                  )}
                </div>

                <div className="grid grid-cols-4 gap-2 md:gap-3">
                  {BREAKFAST_FIELDS.map((field) => (
                    <div key={field.key} className="flex flex-col items-center">
                      <label className="text-[10px] md:text-xs text-slate-400 mb-1 text-center">
                        {field.label}
                      </label>
                      <input
                        type="number"
                        min="0"
                        inputMode="numeric"
                        value={row[field.key as FieldKey]}
                        onChange={(e) =>
                          updateField(row.id, field.key, e.target.value)
                        }
                        placeholder="0"
                        className="w-full text-center rounded-lg md:rounded-xl border border-slate-200 px-1 py-2 md:px-3 md:py-2.5 text-sm md:text-base outline-none focus:ring-2 focus:ring-slate-400 focus:border-slate-400"
                      />
                      <span className="text-[10px] md:text-xs text-slate-300 mt-0.5">
                        {field.unit}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>

        {/* Footer */}
        <div className="border-t border-slate-100 p-4 md:p-6 bg-slate-50">
          <div className="text-xs md:text-sm text-slate-400 text-center space-y-1">
            <div>每天中午12:00自动归档 · 历史归档 {history.length} 条</div>
            <div className="md:hidden">下拉刷新或重新打开链接即可更新</div>
          </div>
        </div>
      </div>
    </div>
  );
}
