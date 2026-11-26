'use client';

import React, { useState, useEffect, createContext, useContext } from 'react';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  Theme,
  IdPrefixProvider,
} from '@fluentui/react-components';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';

// TASK-29: Tema Escuro Corporativo Refinado (Cinza Escuro Suave, sem preto puro #000000)
export const customDarkTheme: Theme = {
  ...webDarkTheme,
  colorNeutralBackground1: '#1f1f1f',
  colorNeutralBackground1Hover: '#2a2a2a',
  colorNeutralBackground1Pressed: '#303030',
  colorNeutralBackground1Selected: '#333333',
  colorNeutralBackground2: '#252526',
  colorNeutralBackground2Hover: '#2f2f30',
  colorNeutralBackground3: '#2d2d2d',
  colorNeutralBackground3Hover: '#363636',
  colorNeutralBackground4: '#333333',
  colorNeutralBackground5: '#383838',
  colorNeutralBackground6: '#3d3d3d',
  colorSubtleBackground: 'transparent',
  colorSubtleBackgroundHover: '#2d2d2d',
  colorSubtleBackgroundPressed: '#252526',
  colorSubtleBackgroundSelected: '#333333',
  colorNeutralCardBackground: '#252526',
  colorNeutralCardBackgroundHover: '#2d2d2d',
  colorNeutralCardBackgroundPressed: '#202020',
  colorNeutralCardBackgroundSelected: '#333333',
  colorNeutralStroke1: '#3d3d3d',
  colorNeutralStroke1Hover: '#4d4d4d',
  colorNeutralStroke1Pressed: '#575757',
  colorNeutralStroke1Selected: '#616161',
  colorNeutralStroke2: '#383838',
  colorNeutralStroke3: '#323232',
  colorNeutralStrokeSubtle: '#2d2d2d',
  colorNeutralForeground1: '#f3f3f3',
  colorNeutralForeground1Hover: '#ffffff',
  colorNeutralForeground1Pressed: '#ffffff',
  colorNeutralForeground1Selected: '#ffffff',
  colorNeutralForeground2: '#d6d6d6',
  colorNeutralForeground2Hover: '#f0f0f0',
  colorNeutralForeground3: '#b3b3b3',
  colorNeutralForeground4: '#8a8886',
};

interface ThemeContextType {
  isDark: boolean;
  toggleTheme: () => void;
}

const ThemeContext = createContext<ThemeContextType>({
  isDark: false,
  toggleTheme: () => {},
});

export const useAppTheme = () => useContext(ThemeContext);

export function Providers({ children }: { children: React.ReactNode }) {
  const [isDark, setIsDark] = useState(false);

  useEffect(() => {
    if (typeof window !== 'undefined') {
      const saved = localStorage.getItem('sp_theme');
      if (saved === 'dark') setIsDark(true);

      // TASK-30: Limpar dados residuais antigos do localStorage e sessionStorage
      try {
        localStorage.removeItem('sp_token');
        localStorage.removeItem('sp_user');
        sessionStorage.clear();
      } catch {}

      // Neutralizar logs de console no browser em execução
      const noop = () => {};
      console.log = noop;
      console.debug = noop;
      console.info = noop;
      console.warn = noop;
    }
  }, []);

  useEffect(() => {
    if (typeof document !== 'undefined') {
      document.body.style.backgroundColor = isDark ? '#1f1f1f' : '#ffffff';
      document.body.style.color = isDark ? '#f3f3f3' : '#242424';
    }
  }, [isDark]);

  const toggleTheme = () => {
    setIsDark((prev) => {
      const next = !prev;
      if (typeof window !== 'undefined') {
        localStorage.setItem('sp_theme', next ? 'dark' : 'light');
      }
      return next;
    });
  };

  const [queryClient] = useState(
    () =>
      new QueryClient({
        defaultOptions: {
          queries: {
            staleTime: 1000 * 60 * 5,
            refetchOnWindowFocus: false,
          },
        },
      })
  );

  return (
    <ThemeContext.Provider value={{ isDark, toggleTheme }}>
      <IdPrefixProvider value="sp-">
        <FluentProvider theme={isDark ? customDarkTheme : webLightTheme}>
          <QueryClientProvider client={queryClient}>
            {children}
          </QueryClientProvider>
        </FluentProvider>
      </IdPrefixProvider>
    </ThemeContext.Provider>
  );
}
