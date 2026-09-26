import { randomUUID } from 'node:crypto';

export type Market = 'TWSE' | 'TPEX';

export interface Company {
  id: string;
  market: Market;
  stockCode: string;
  name: string;
  watchlist: {
    active: boolean;
    category: string;
    notes: string;
  };
}

export function createCompany(input: { market: Market; stockCode: string; name: string }): Company {
  const stockCode = input.stockCode.trim();
  const name = input.name.trim();

  if (input.market !== 'TWSE' && input.market !== 'TPEX') {
    throw new Error('市場必須是 TWSE 或 TPEX');
  }
  if (!/^\d{4,6}$/.test(stockCode)) {
    throw new Error('股票代號必須是 4 至 6 位數字');
  }
  if (!name) {
    throw new Error('公司名稱不可空白');
  }

  return {
    id: randomUUID(),
    market: input.market,
    stockCode,
    name,
    watchlist: { active: true, category: '', notes: '' },
  };
}

export function setWatchlistActive(company: Company, active: boolean): Company {
  return { ...company, watchlist: { ...company.watchlist, active } };
}

export function updateWatchlistDetails(company: Company, details: { category: string; notes: string }): Company {
  return {
    ...company,
    watchlist: {
      ...company.watchlist,
      category: details.category.trim(),
      notes: details.notes.trim(),
    },
  };
}
