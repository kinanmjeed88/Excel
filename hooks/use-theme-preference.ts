import AsyncStorage from "@react-native-async-storage/async-storage";
import { useCallback, useEffect, useMemo, useState } from "react";

import { Colors, type ColorScheme, type ThemeColorPalette } from "@/constants/theme";
import { useColors } from "./use-colors";
import { useThemeContext } from "@/lib/theme-provider";

const THEME_PREFERENCE_KEY = "jadwali.colorScheme.v1";

type ThemePreference = {
  /** الوضع النشط حالياً في التطبيق (يدوياً أو تبعاً للنظام). */
  activeScheme: ColorScheme;
  /** تفضيل المستخدم المحفوظ، أو "system" لتبع النظام تلقائياً. */
  savedPreference: ColorScheme | "system";
  /** تبديل الوضع يدوياً (light/dark) مع حفظ التفضيل. */
  toggleColorScheme: (scheme: ColorScheme) => void;
  /** هل الوضع الداكن نشط حالياً. */
  isDark: boolean;
  /** لوحة الألوان النشطة من السمة المركزية. */
  colors: ThemeColorPalette;
};

/**
 * يدير وضع العرض (فاتح/داكن) مع حفظ تفضيل المستخدم في AsyncStorage.
 * عند "system" يتبع إعداد النظام تلقائياً.
 */
export function useThemePreference(): ThemePreference {
  const { colorScheme, setColorScheme } = useThemeContext();
  const colors = useColors();
  const [savedPreference, setSavedPreference] = useState<ColorScheme | "system">("system");

  useEffect(() => {
    let isActive = true;
    AsyncStorage.getItem(THEME_PREFERENCE_KEY)
      .then((stored) => {
        if (!stored) return;
        const parsed = stored as ColorScheme | "system";
        if (parsed === "system" || parsed === "light" || parsed === "dark") {
          if (isActive) setSavedPreference(parsed);
          if (parsed !== "system") setColorScheme(parsed);
        }
      })
      .catch(() => undefined);
    return () => {
      isActive = false;
    };
  }, [setColorScheme]);

  const toggleColorScheme = useCallback(
    (scheme: ColorScheme) => {
      setSavedPreference(scheme);
      setColorScheme(scheme);
      AsyncStorage.setItem(THEME_PREFERENCE_KEY, scheme).catch(() => undefined);
    },
    [setColorScheme],
  );

  return useMemo(
    () => ({
      activeScheme: colorScheme,
      savedPreference,
      toggleColorScheme,
      isDark: colorScheme === "dark",
      colors,
    }),
    [colorScheme, savedPreference, toggleColorScheme, colors],
  );
}

