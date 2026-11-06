import { createContext } from "react";
export const GuideContext = createContext<{
  openPage: () => void;
  openIndex: () => void;
  openSetup?: () => void;
  startQuick: () => void;
  skip: () => void;
  welcome: boolean;
  message: string | null;
} | null>(null);
