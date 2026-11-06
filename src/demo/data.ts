export type Page =
  | "總覽"
  | "關注公司"
  | "重大訊息"
  | "違約交割"
  | "法說會"
  | "盤面紀錄"
  | "設定";
export const pages: Page[] = [
  "總覽",
  "關注公司",
  "重大訊息",
  "違約交割",
  "法說會",
  "盤面紀錄",
  "設定",
];
export interface Company {
  code: string;
  name: string;
  market: string;
  category: string;
  active: boolean;
  notes: string;
}
export const companies: Company[] = [
  {
    code: "2330",
    name: "台積電",
    market: "上市",
    category: "半導體",
    active: true,
    notes: "",
  },
  {
    code: "2454",
    name: "聯發科",
    market: "上市",
    category: "半導體",
    active: true,
    notes: "",
  },
  {
    code: "2317",
    name: "鴻海",
    market: "上市",
    category: "電子",
    active: true,
    notes: "",
  },
  {
    code: "3105",
    name: "穩懋",
    market: "上櫃",
    category: "半導體",
    active: false,
    notes: "",
  },
  {
    code: "2308",
    name: "台達電",
    market: "上市",
    category: "電子",
    active: true,
    notes: "",
  },
  {
    code: "3711",
    name: "日月光投控",
    market: "上市",
    category: "半導體",
    active: true,
    notes: "",
  },
  {
    code: "2303",
    name: "聯電",
    market: "上市",
    category: "半導體",
    active: true,
    notes: "",
  },
  {
    code: "2412",
    name: "中華電",
    market: "上市",
    category: "電信",
    active: true,
    notes: "",
  },
];
export interface Announcement {
  id: number;
  time: string;
  company: string;
  code: string;
  title: string;
  body: string;
  unread: boolean;
  watched: boolean;
}
export const announcements: Announcement[] = [
  {
    id: 1,
    time: "11/05 14:32",
    company: "台積電",
    code: "2330",
    title: "董事會決議股利分配",
    body: "公告董事會決議股利分配事宜。詳細內容請以公開資訊觀測站公告為準。",
    unread: true,
    watched: true,
  },
  {
    id: 2,
    time: "11/05 14:10",
    company: "聯發科",
    code: "2454",
    title: "公告召開法人說明會",
    body: "公司將召開法人說明會，說明近期營運概況與財務資訊。",
    unread: true,
    watched: true,
  },
  {
    id: 3,
    time: "11/05 13:45",
    company: "鴻海",
    code: "2317",
    title: "董事會重要決議",
    body: "董事會重要決議事項摘要，實際資訊以官方公告為準。",
    unread: false,
    watched: true,
  },
  {
    id: 4,
    time: "11/05 13:20",
    company: "穩懋",
    code: "3105",
    title: "公告營運概況",
    body: "公司近期營運概況說明。本內容僅供介面示範。",
    unread: false,
    watched: true,
  },
  {
    id: 5,
    time: "11/05 12:40",
    company: "台達電",
    code: "2308",
    title: "公告董事會召開日期",
    body: "本內容僅供介面示範。",
    unread: true,
    watched: true,
  },
  {
    id: 6,
    time: "11/05 11:05",
    company: "聯電",
    code: "2303",
    title: "公告本公司相關資訊",
    body: "本內容僅供介面示範。",
    unread: true,
    watched: true,
  },
  {
    id: 7,
    time: "11/05 10:15",
    company: "中華電",
    code: "2412",
    title: "公告法說會相關資訊",
    body: "本內容僅供介面示範。",
    unread: true,
    watched: true,
  },
];
export interface Preferences {
  material: boolean;
  start: string;
  end: string;
  frequency: string;
  disclosure: boolean;
  disclosureTime: string;
  notifyEmpty: boolean;
  login: boolean;
  hidden: boolean;
  notifications: boolean;
}
export const initialPreferences: Preferences = {
  material: true,
  start: "09:00",
  end: "18:00",
  frequency: "15",
  disclosure: true,
  disclosureTime: "18:30",
  notifyEmpty: true,
  login: true,
  hidden: true,
  notifications: true,
};
export const conferences = [
  {
    date: "11/12",
    day: "週四",
    time: "14:00–16:00",
    name: "台積電",
    code: "2330",
    place: "線上法說會",
  },
  {
    date: "11/18",
    day: "週三",
    time: "14:30–16:30",
    name: "聯發科",
    code: "2454",
    place: "線上直播",
  },
  {
    date: "11/24",
    day: "週二",
    time: "15:00–17:00",
    name: "鴻海",
    code: "2317",
    place: "公司說明會",
  },
];
export const indexRanks = {
  rise: [
    ["2026/07/31", "3,281.06"],
    ["2026/05/04", "1,828.89"],
    ["2026/07/21", "1,783.17"],
  ],
  fall: [
    ["2026/07/17", "2,953.71"],
    ["2026/06/08", "2,694.08"],
    ["2026/07/29", "2,218.51"],
  ],
};
export const investorRanks: Record<
  string,
  {
    days: string;
    amount: string;
    rank: string;
    buy: string[][];
    sell: string[][];
  }
> = {
  投信: {
    days: "17",
    amount: "3,044.45",
    rank: "2",
    buy: [
      ["2026/06/23", "2026/08/06", "32", "3,809.89"],
      ["2026/10/08", "2026/11/04", "17", "3,044.45"],
      ["2024/03/01", "2024/04/09", "26", "2,163.01"],
    ],
    sell: [
      ["2025/12/24", "2026/01/30", "26", "1,367.42"],
      ["2025/08/26", "2025/09/26", "24", "845.53"],
      ["2026/04/14", "2026/04/20", "5", "280.68"],
    ],
  },
  外資: {
    days: "5",
    amount: "5,120.10",
    rank: "3",
    buy: [
      ["2026/05/21", "2026/05/27", "5", "2,595.01"],
      ["2026/08/10", "2026/08/17", "6", "2,515.83"],
      ["2023/01/05", "2023/02/03", "14", "2,503.49"],
    ],
    sell: [
      ["2026/10/06", "2026/10/20", "10", "8,380.65"],
      ["2026/07/02", "2026/07/21", "13", "6,521.99"],
      ["2026/10/29", "2026/11/04", "5", "5,120.10"],
    ],
  },
  自營商: {
    days: "5",
    amount: "1,847.31",
    rank: "3",
    buy: [
      ["2025/07/15", "2025/08/01", "14", "765.97"],
      ["2025/12/17", "2025/12/30", "9", "749.87"],
      ["2026/09/17", "2026/09/24", "6", "716.28"],
    ],
    sell: [
      ["2026/10/07", "2026/10/27", "12", "3,636.96"],
      ["2024/07/05", "2024/08/15", "28", "3,189.81"],
      ["2026/10/29", "2026/11/04", "5", "1,847.31"],
    ],
  },
};
