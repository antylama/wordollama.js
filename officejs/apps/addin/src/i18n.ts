import i18n from "i18next";
import { initReactI18next } from "react-i18next";
import enUS from "./settings/locales/en-US.json" with { type: "json" };
import plPL from "./settings/locales/pl-PL.json" with { type: "json" };
import zhCN from "./settings/locales/zh-CN.json" with { type: "json" };

export const UI_LOCALE_STORAGE_KEY = "wordollama-ui-locale";
export type UiLocalePreference = "auto" | "en-US" | "zh-CN" | "pl-PL";
export type SupportedUiLocale = Exclude<UiLocalePreference, "auto">;

function normalizeLocale(locale: string | undefined): SupportedUiLocale {
  const normalized = locale?.toLowerCase() ?? "";
  if (normalized.startsWith("zh")) return "zh-CN";
  if (normalized.startsWith("pl")) return "pl-PL";
  return "en-US";
}

export function readUiLocalePreference(): UiLocalePreference {
  if (typeof window === "undefined") return "auto";
  const stored = window.localStorage.getItem(UI_LOCALE_STORAGE_KEY);
  return stored === "en-US" || stored === "zh-CN" || stored === "pl-PL" ? stored : "auto";
}

export function resolveOfficeLocale(): SupportedUiLocale {
  const officeLocale = typeof Office !== "undefined"
    ? String(Office.context?.displayLanguage ?? "")
    : "";
  const browserLocale = typeof navigator !== "undefined" ? navigator.language : "en-US";
  return normalizeLocale(officeLocale || browserLocale);
}

export function resolveUiLocale(preference = readUiLocalePreference()): SupportedUiLocale {
  return preference === "auto" ? resolveOfficeLocale() : preference;
}

export function getCurrentUiLocale(): SupportedUiLocale {
  return normalizeLocale(i18n.resolvedLanguage ?? resolveUiLocale());
}

void i18n
  .use(initReactI18next)
  .init({
    resources: {
      "en-US": { translation: enUS },
      "zh-CN": { translation: zhCN },
      "pl-PL": { translation: plPL },
    },
    lng: resolveUiLocale(),
    fallbackLng: "en-US",
    supportedLngs: ["en-US", "zh-CN", "pl-PL"],
    interpolation: { escapeValue: false },
    returnNull: false,
    initAsync: false,
  });

export async function setUiLocalePreference(preference: UiLocalePreference): Promise<void> {
  if (typeof window !== "undefined") {
    if (preference === "auto") {
      window.localStorage.removeItem(UI_LOCALE_STORAGE_KEY);
    } else {
      window.localStorage.setItem(UI_LOCALE_STORAGE_KEY, preference);
    }
  }
  const locale = resolveUiLocale(preference);
  if (typeof document !== "undefined") document.documentElement.lang = locale;
  await i18n.changeLanguage(locale);
}

if (typeof document !== "undefined") document.documentElement.lang = resolveUiLocale();

export default i18n;
