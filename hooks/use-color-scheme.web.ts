import { useEffect, useState } from "react";
import { useColorScheme as useRNColorScheme } from "react-native";

/**
 * To support static rendering, this value needs to be re-calculated on the client side for web
 */
export function useColorScheme() {
  const [hasHydrated, setHasHydrated] = useState(false);

  useEffect(() => {
    setHasHydrated(true);
  }, []);

  // على الويب، نتبع وضع العرض المختار يدوياً (dataset.theme) أو النظام.
  const rootTheme =
    typeof document !== "undefined"
      ? (document.documentElement.dataset.theme as "light" | "dark" | undefined)
      : undefined;

  const systemScheme = useRNColorScheme();
  const colorScheme = rootTheme === "dark" ? ("dark" as const) : (systemScheme ?? "light");

  if (hasHydrated) {
    return colorScheme;
  }

  return rootTheme ?? "light";
}
