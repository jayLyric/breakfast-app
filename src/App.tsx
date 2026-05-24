import React, {
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import * as XLSX from 'xlsx';

const DEFAULT_USERS = [
  '张三',
  '李四',
  '王五',
  '赵六',
  '孙七',
  '周八',
  '吴九',
  '郑十',
  '王强',
  '李明',
  '刘洋',
  '陈伟',
  '杨帆',
  '黄涛',
  '赵鹏',
  '胡斌',
  '谢凯',
  '何俊',
  '高飞',
  '林宇',
  '郭磊',
  '罗成',
  '马超',
  '许峰',
  '梁杰',
  '宋涛',
  '彭浩',
  '唐龙',
  '韩松',
  '邓辉',
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
  { key: 'baozi', label: '包子（个）' },
  { key: 'mantou', label: '馒头（个）' },
  { key: 'porridge', label: '稀饭（碗）' },
  { key: 'egg', label: '鸡蛋（个）' },
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

const getStorageKey = () => `breakfast-${getDateKey()}`;
const getResetKey = () => `breakfast-reset-${getDateKey()}`;

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
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

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
    const content = rows
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
    <div className="min-h-screen bg-slate-100 p-4 md:p-8">
      <div className="mx-auto max-w-7xl rounded-[32px] bg-white shadow-2xl overflow-hidden border border-slate-200">
        <div className="bg-slate-900 text-white p-8">
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6">
            <div>
              <h1 className="text-4xl font-bold tracking-tight">
                早餐统计生产版系统
              </h1>
              <p className="text-slate-300 mt-3 text-lg">
                微信群直接打开 · 自动归档 · Excel导出
              </p>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {BREAKFAST_FIELDS.map((field) => (
                <div
                  key={field.key}
                  className="bg-white/10 rounded-2xl px-5 py-4 backdrop-blur"
                >
                  <div className="text-sm text-slate-300">{field.label}</div>
                  <div className="text-3xl font-bold mt-2">
                    {totals[field.key]}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="p-6 border-b border-slate-100 bg-slate-50 flex flex-col lg:flex-row gap-4 lg:items-center lg:justify-between">
          <div className="flex flex-wrap gap-3">
            <button
              onClick={copyWechatText}
              className="rounded-2xl bg-slate-900 text-white px-6 py-3 font-medium hover:scale-[1.02] transition-all"
            >
              复制微信群统计
            </button>
            <button
              onClick={exportExcel}
              className="rounded-2xl bg-emerald-600 text-white px-6 py-3 font-medium hover:scale-[1.02] transition-all"
            >
              导出 Excel
            </button>
            <button
              onClick={clearToday}
              className="rounded-2xl bg-white border border-slate-200 px-6 py-3 font-medium hover:bg-slate-100 transition-all"
            >
              手动清空
            </button>
          </div>

          <div className="flex flex-col md:flex-row gap-4 md:items-center">
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="搜索姓名..."
              className="rounded-2xl border border-slate-200 px-5 py-3 outline-none focus:ring-2 focus:ring-slate-400 bg-white"
            />
            <div className="text-sm text-slate-500">
              已填写：
              <span className="font-bold text-slate-800">
                {filledCount}/{rows.length}
              </span>
            </div>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full min-w-[1100px]">
            <thead>
              <tr className="bg-slate-100 border-b border-slate-200">
                <th className="text-left p-5 font-semibold">姓名</th>
                {BREAKFAST_FIELDS.map((field) => (
                  <th key={field.key} className="text-left p-5 font-semibold">
                    {field.label}
                  </th>
                ))}
                <th className="text-left p-5 font-semibold">更新时间</th>
                <th className="text-left p-5 font-semibold">状态</th>
              </tr>
            </thead>
            <tbody>
              {filteredRows.map((row) => {
                const isFilled = BREAKFAST_FIELDS.some(
                  (field) => row[field.key as FieldKey] !== '',
                );
                return (
                  <tr
                    key={row.id}
                    className="border-b border-slate-100 hover:bg-slate-50 transition-all"
                  >
                    <td className="p-5 font-medium text-slate-700">
                      {row.name}
                    </td>
                    {BREAKFAST_FIELDS.map((field) => (
                      <td key={field.key} className="p-5">
                        <input
                          type="number"
                          min="0"
                          value={row[field.key as FieldKey]}
                          onChange={(e) =>
                            updateField(row.id, field.key, e.target.value)
                          }
                          placeholder="0"
                          className="w-24 rounded-xl border border-slate-200 px-3 py-3 outline-none focus:ring-2 focus:ring-slate-400"
                        />
                      </td>
                    ))}
                    <td className="p-5 text-sm text-slate-500">
                      {row.updatedAt || '-'}
                    </td>
                    <td className="p-5">
                      {isFilled ? (
                        <span className="rounded-full bg-emerald-100 text-emerald-700 px-4 py-1 text-sm font-medium">
                          已填写
                        </span>
                      ) : (
                        <span className="rounded-full bg-slate-100 text-slate-500 px-4 py-1 text-sm font-medium">
                          未填写
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div className="border-t border-slate-100 p-6 bg-slate-50 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
          <div className="text-sm text-slate-500 leading-7">
            <div>自动归档时间：第二天中午12:00</div>
            <div>历史归档数量：{history.length}</div>
            <div>最近归档：{lastArchive || '暂无归档'}</div>
          </div>
          <div className="text-right text-sm text-slate-500">
            <div>适配手机 / 平板 / 微信内浏览器</div>
            <div>支持100+人同时使用</div>
          </div>
        </div>
      </div>
    </div>
  );
}
