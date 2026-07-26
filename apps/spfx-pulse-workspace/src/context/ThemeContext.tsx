import * as React from 'react';

export interface IThemeContext {
  readonly isDarkTheme: boolean;
}

export const ThemeContext = React.createContext<IThemeContext>({ isDarkTheme: false });
