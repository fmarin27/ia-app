import React, {
  startTransition,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import AsyncStorage from "@react-native-async-storage/async-storage";
import {
  ActivityIndicator,
  Alert,
  FlatList,
  Image,
  Linking,
  Modal,
  Platform,
  Pressable,
  SafeAreaView,
  ScrollView,
  StyleSheet,
  StatusBar as NativeStatusBar,
  Text,
  TextInput,
  Vibration,
  View,
} from "react-native";
import { StatusBar as ExpoStatusBar } from "expo-status-bar";
import { CameraView, useCameraPermissions } from "expo-camera";
import * as DocumentPicker from "expo-document-picker";
import * as FileSystem from "expo-file-system/legacy";
import * as ImagePicker from "expo-image-picker";
import * as ImageManipulator from "expo-image-manipulator";
import * as Updates from "expo-updates";
import appConfig from "./app.json";

function normalizeBaseUrl(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}

function cleanString(value) {
  return String(value || "").trim();
}

function parseFallbackBases(value) {
  return String(value || "")
    .split(",")
    .map((entry) => normalizeBaseUrl(entry))
    .filter((entry) => entry && entry !== "__NONE__");
}

function normalizeVariantKey(value) {
  return cleanString(value).toLowerCase().replace(/[^a-z0-9]+/g, "_") || "default";
}

function parseBooleanFlag(value) {
  return cleanString(value).toLowerCase() === "true";
}

function parseUiScaleValue(value, fallback = 1) {
  const parsed = Number.parseFloat(String(value ?? fallback));
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

const UPDATE_CHANNEL = cleanString(Updates.channel).toLowerCase();
const CHANNEL_PRESETS = {
  joe: {
    appVariant: "joe",
    ownerName: "Joe Lasala",
    apiBase: "https://joe-api.luxuryimportsusa.shop",
    apiFallbacks: [],
    disableFontScaling: true,
    uiScale: 0.9,
    lockConnection: true,
  },
  production: {
    appVariant: "default",
    ownerName: "Fernando Marin",
    apiBase: "https://api.luxuryimportsusa.shop",
    apiFallbacks: ["https://api2.luxuryimportsusa.shop"],
    disableFontScaling: false,
    uiScale: 1,
    lockConnection: false,
  },
};
const ACTIVE_CHANNEL_PRESET = CHANNEL_PRESETS[UPDATE_CHANNEL] || null;
const API_BASE_DEFAULT =
  normalizeBaseUrl(ACTIVE_CHANNEL_PRESET?.apiBase || process.env.EXPO_PUBLIC_API_BASE) ||
  "https://api.luxuryimportsusa.shop";
const API_BASE_FALLBACKS = parseFallbackBases(
  ACTIVE_CHANNEL_PRESET ? ACTIVE_CHANNEL_PRESET.apiFallbacks.join(",") : process.env.EXPO_PUBLIC_API_FALLBACKS ?? "https://api2.luxuryimportsusa.shop"
).filter((entry, index, values) => entry !== API_BASE_DEFAULT && values.indexOf(entry) === index);
const APP_VARIANT = normalizeVariantKey(ACTIVE_CHANNEL_PRESET?.appVariant || process.env.EXPO_PUBLIC_APP_VARIANT || "default");
const APP_OWNER_LABEL =
  cleanString(ACTIVE_CHANNEL_PRESET?.ownerName || process.env.EXPO_PUBLIC_APP_OWNER_NAME) || "Fernando Marin";
const DISABLE_FONT_SCALING = ACTIVE_CHANNEL_PRESET
  ? Boolean(ACTIVE_CHANNEL_PRESET.disableFontScaling)
  : parseBooleanFlag(process.env.EXPO_PUBLIC_DISABLE_FONT_SCALING);
const UI_SCALE = ACTIVE_CHANNEL_PRESET
  ? parseUiScaleValue(ACTIVE_CHANNEL_PRESET.uiScale, 1)
  : parseUiScaleValue(process.env.EXPO_PUBLIC_UI_SCALE, 1);
const CONNECTION_LOCKED = ACTIVE_CHANNEL_PRESET
  ? Boolean(ACTIVE_CHANNEL_PRESET.lockConnection)
  : parseBooleanFlag(process.env.EXPO_PUBLIC_LOCK_CONNECTION);
const BULLET = " | ";
const API_BASE_STORAGE_KEY = `claim_manager_mobile_api_base_${APP_VARIANT}`;
const ROUTE_PLAN_STORAGE_KEY = `claim_manager_mobile_route_plan_keys_${APP_VARIANT}`;
const ROUTE_ADDRESS_OVERRIDES_STORAGE_KEY = `claim_manager_mobile_route_address_overrides_${APP_VARIANT}`;
const PHOTO_UPLOAD_QUEUE_STORAGE_KEY = `claim_manager_mobile_photo_upload_queue_${APP_VARIANT}`;
const APP_VERSION = appConfig?.expo?.version || "unknown";
const APP_RELEASE_LABEL =
  cleanString(process.env.EXPO_PUBLIC_RELEASE_LABEL) ||
  cleanString(appConfig?.expo?.extra?.releaseLabel) ||
  `${UPDATE_CHANNEL || "default"}-runtime`;
const CONNECTION_HINT = CONNECTION_LOCKED
  ? "This build is pinned to Joe's PC endpoint."
  : "Use the secure Cloudflare address by default, or switch back to local if you are on the same network.";

function scaleUi(value) {
  return Math.round(value * UI_SCALE * 100) / 100;
}

function hostnameFromBase(value) {
  try {
    return new URL(normalizeBaseUrl(value)).hostname.toLowerCase();
  } catch {
    return "";
  }
}

function isLocalNetworkBase(value) {
  const host = hostnameFromBase(value);
  if (!host) {
    return false;
  }
  if (host === "localhost" || host === "127.0.0.1") {
    return true;
  }
  if (/^192\.168\./.test(host) || /^10\./.test(host)) {
    return true;
  }
  const match172 = host.match(/^172\.(\d{1,3})\./);
  if (match172) {
    const second = Number.parseInt(match172[1], 10);
    if (second >= 16 && second <= 31) {
      return true;
    }
  }
  return false;
}

function isLegacyOrDisallowedPublicBase(value) {
  const normalized = normalizeBaseUrl(value);
  if (!normalized) {
    return false;
  }
  if (normalized === API_BASE_DEFAULT) {
    return false;
  }
  if (API_BASE_FALLBACKS.includes(normalized)) {
    return true;
  }
  return !isLocalNetworkBase(normalized);
}

if (DISABLE_FONT_SCALING) {
  Text.defaultProps = {
    ...(Text.defaultProps || {}),
    allowFontScaling: false,
    maxFontSizeMultiplier: 1,
  };
  TextInput.defaultProps = {
    ...(TextInput.defaultProps || {}),
    allowFontScaling: false,
    maxFontSizeMultiplier: 1,
  };
}

function digitsOnlyPhone(value) {
  return String(value || "").replace(/[^\d+]/g, "");
}

function formatPhone(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (digits.length === 11 && digits.startsWith("1")) {
    return `+1 (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7)}`;
  }
  if (digits.length === 10) {
    return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
  }
  return String(value || "").trim();
}

const HOME_TABS = [
  ["claims", "Claims"],
  ["route", "Route Planner"],
  ["library", "Library"],
];

const CLAIM_TABS = [
  ["open", "Open"],
  ["closed", "Closed"],
  ["all", "All"],
  ["new", "New"],
];

const ROUTE_STATUS_TABS = [
  ["all", "All Claims"],
  ["open", "Open"],
  ["closed", "Closed"],
];

const LIBRARY_TABS = [
  ["tools", "Claim Tools"],
  ["shops", "Body Shops"],
  ["insurance", "Insurance"],
];

const PHOTO_LABEL_OPTIONS = [
  "VIN",
  "MILEAGE",
  "MEASUREMENT",
  "LICENSE PLATE",
  "INTERIOR",
  "DASHBOARD",
  "DOOR TRIM PANEL",
  "LEFT FRONT",
  "RIGHT FRONT",
  "LEFT REAR",
  "RIGHT REAR",
  "DAMAGE",
  "UPD",
  "TIRE INFO",
  "INVOICE",
  "TOW BILL",
  "ESTIMATE",
  "DOP",
  "REPAIR AUTHORIZATION",
  "OTHER",
];

const COMMON_EMAIL_RECIPIENTS = [
  { label: "Joe", email: "joe@lasalallc.com" },
  { label: "Lisa", email: "ldellacorte@duhamels.com" },
  { label: "Donna", email: "dnoone@duhamels.com" },
  { label: "Gina", email: "gferreira@duhamels.com" },
];

const PHOTO_ZOOM_OPTIONS = [
  { label: "1x", value: 0 },
  { label: "1.5x", value: 0.15 },
  { label: "2x", value: 0.3 },
  { label: "3x", value: 0.5 },
];

const MOBILE_PHOTO_MAX_SIDE = 1920;
const MOBILE_PHOTO_PRIMARY_COMPRESSION = 0.62;
const MOBILE_PHOTO_SECONDARY_COMPRESSION = 0.54;
const MOBILE_PHOTO_HEAVY_COMPRESSION = 0.46;
const PHOTO_UPLOAD_QUEUE_RETRY_MS = 30000;
const PHOTO_UPLOAD_QUEUE_DIR = `${FileSystem.documentDirectory || FileSystem.cacheDirectory || ""}claim-photo-upload-queue/`;

function sanitizePhotoQueueName(value) {
  const cleaned = cleanString(value).replace(/\.[^.]+$/, "");
  const safe = cleaned.replace(/[^a-z0-9_-]+/gi, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
  return safe || "claim-photo";
}

function createPhotoQueueId() {
  return `photo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function queueUploadLabel(item) {
  return cleanString(item?.label) || cleanString(item?.name) || "Photo";
}

async function prepareUploadPhoto(asset) {
  const manipulations = [];
  const width = asset.width || 0;
  const height = asset.height || 0;
  const longestSide = Math.max(width, height);

  if (longestSide > MOBILE_PHOTO_MAX_SIDE) {
    if (width >= height) {
      manipulations.push({ resize: { width: MOBILE_PHOTO_MAX_SIDE } });
    } else {
      manipulations.push({ resize: { height: MOBILE_PHOTO_MAX_SIDE } });
    }
  }

  let compression = MOBILE_PHOTO_PRIMARY_COMPRESSION;
  if ((asset.fileSize || 0) > 3000000 || longestSide > 2500) {
    compression = MOBILE_PHOTO_SECONDARY_COMPRESSION;
  }
  if ((asset.fileSize || 0) > 6000000 || longestSide > 3500) {
    compression = MOBILE_PHOTO_HEAVY_COMPRESSION;
  }

  let result = await ImageManipulator.manipulateAsync(
    asset.uri,
    manipulations,
    {
      compress: compression,
      format: ImageManipulator.SaveFormat.JPEG,
      base64: false,
    }
  );

  if ((asset.fileSize || 0) > 8000000) {
    result = await ImageManipulator.manipulateAsync(
      result.uri,
      [],
      {
        compress: MOBILE_PHOTO_HEAVY_COMPRESSION,
        format: ImageManipulator.SaveFormat.JPEG,
        base64: false,
      }
    );
  }

  return {
    uri: result.uri,
    name: (asset.fileName || "claim-photo").replace(/\.[^.]+$/, "") + ".jpg",
    type: "image/jpeg",
    capturedAt:
      asset?.capturedAt ||
      asset?.creationTime ||
      asset?.modificationTime ||
      new Date().toISOString(),
  };
}

export default function App() {
  const [apiBase, setApiBase] = useState(API_BASE_DEFAULT);
  const [utilityOpen, setUtilityOpen] = useState(false);
  const [homeTab, setHomeTab] = useState("claims");
  const [claimTab, setClaimTab] = useState("open");
  const [routeStatus, setRouteStatus] = useState("open");
  const [libraryTab, setLibraryTab] = useState("tools");
  const [search, setSearch] = useState("");
  const [dashboard, setDashboard] = useState({
    summary: { all: 0, open: 0, closed: 0 },
    settings: {},
  });
  const [claimRecords, setClaimRecords] = useState([]);
  const [claimPhoneMap, setClaimPhoneMap] = useState({});
  const [routeData, setRouteData] = useState({
    start_from: "",
    records: [],
    planned_records: [],
  });
  const [routePlanKeys, setRoutePlanKeys] = useState([]);
  const [toolsData, setToolsData] = useState({
    claim_tools_folder: "",
    contacts: [],
    files: [],
  });
  const [bodyShops, setBodyShops] = useState([]);
  const [insuranceCompanies, setInsuranceCompanies] = useState([]);
  const [selectedClaimKey, setSelectedClaimKey] = useState("");
  const [claimDetail, setClaimDetail] = useState(null);
  const [detailVisible, setDetailVisible] = useState(false);
  const [loading, setLoading] = useState(false);
  const [detailLoading, setDetailLoading] = useState(false);
  const [noteSaving, setNoteSaving] = useState(false);
  const [uploadingPhoto, setUploadingPhoto] = useState(false);
  const [emailModalVisible, setEmailModalVisible] = useState(false);
  const [selectedPdfPaths, setSelectedPdfPaths] = useState([]);
  const [emailTo, setEmailTo] = useState("");
  const [emailCc, setEmailCc] = useState("");
  const [emailSubject, setEmailSubject] = useState("");
  const [emailBody, setEmailBody] = useState("");
  const [emailSending, setEmailSending] = useState(false);
  const [emailRecipientPickerOpen, setEmailRecipientPickerOpen] = useState(false);
  const [assignmentUploadInFlight, setAssignmentUploadInFlight] = useState(false);
  const [assignmentUploadResult, setAssignmentUploadResult] = useState(null);
  const [photoSessionVisible, setPhotoSessionVisible] = useState(false);
  const [pendingPhotoAsset, setPendingPhotoAsset] = useState(null);
  const [pendingPhotoLabel, setPendingPhotoLabel] = useState("");
  const [photoSessionUploads, setPhotoSessionUploads] = useState([]);
  const [queuedPhotoUploads, setQueuedPhotoUploads] = useState([]);
  const [queuedPhotoUploadsReady, setQueuedPhotoUploadsReady] = useState(false);
  const [queueFlushing, setQueueFlushing] = useState(false);
  const [cameraReady, setCameraReady] = useState(false);
  const [photoZoom, setPhotoZoom] = useState(0);
  const [photoCaptureCueVisible, setPhotoCaptureCueVisible] = useState(false);
  const [photoCaptureCueText, setPhotoCaptureCueText] = useState("");
  const [routeSaving, setRouteSaving] = useState(false);
  const [routePlanReady, setRoutePlanReady] = useState(false);
  const [routeAddressOverrides, setRouteAddressOverrides] = useState({});
  const [routeAddressOverridesReady, setRouteAddressOverridesReady] = useState(false);
  const [editingRouteKey, setEditingRouteKey] = useState("");
  const [editingRouteAddress, setEditingRouteAddress] = useState("");
  const [routeAddressSaving, setRouteAddressSaving] = useState(false);
  const [updateChecking, setUpdateChecking] = useState(false);
  const [hasLoadedOnce, setHasLoadedOnce] = useState(false);
  const [noteText, setNoteText] = useState("");
  const [error, setError] = useState("");
  const [updateMessage, setUpdateMessage] = useState("");
  const [apiReady, setApiReady] = useState(false);
  const loadSequence = useRef(0);
  const cameraRef = useRef(null);
  const photoCaptureCueTimeoutRef = useRef(null);
  const queuedPhotoUploadsRef = useRef([]);
  const queueFlushInFlightRef = useRef(false);
  const deferredSearch = useDeferredValue(search);
  const [cameraPermission, requestCameraPermission] = useCameraPermissions();
  const currentChannel = UPDATE_CHANNEL || "preview";
  const currentRuntimeVersion = Updates.runtimeVersion || APP_VERSION;
  const currentUpdateId = Updates.updateId ? Updates.updateId.slice(0, 8) : "embedded";
  const currentVersionLabel = `v${APP_VERSION} / r${APP_RELEASE_LABEL} / ${currentChannel} / ${currentUpdateId}`;
  const currentVersionDetail = `Runtime ${currentRuntimeVersion}`;
  const listPerfProps = {
    initialNumToRender: 10,
    maxToRenderPerBatch: 10,
    windowSize: 7,
    removeClippedSubviews: true,
  };
  const currentClaimQueuedUploads = useMemo(
    () => queuedPhotoUploads.filter((item) => item.claimKey === selectedClaimKey),
    [queuedPhotoUploads, selectedClaimKey]
  );

  useEffect(() => {
    return () => {
      if (photoCaptureCueTimeoutRef.current) {
        clearTimeout(photoCaptureCueTimeoutRef.current);
      }
    };
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function loadStoredApiBase() {
      if (CONNECTION_LOCKED) {
        if (!cancelled) {
          setApiBase(API_BASE_DEFAULT);
          setApiReady(true);
        }
        return;
      }
      try {
        const saved = normalizeBaseUrl(await AsyncStorage.getItem(API_BASE_STORAGE_KEY));
        if (!cancelled && saved) {
          if (isLegacyOrDisallowedPublicBase(saved)) {
            await AsyncStorage.setItem(API_BASE_STORAGE_KEY, API_BASE_DEFAULT);
            setApiBase(API_BASE_DEFAULT);
          } else {
            setApiBase(saved);
          }
        }
      } catch {}
      if (!cancelled) {
        setApiReady(true);
      }
    }

    loadStoredApiBase();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!apiReady) {
      return;
    }
    if (CONNECTION_LOCKED) {
      return;
    }
    const normalizedBase = normalizeBaseUrl(apiBase);
    if (isLegacyOrDisallowedPublicBase(normalizedBase)) {
      setApiBase(API_BASE_DEFAULT);
      AsyncStorage.setItem(API_BASE_STORAGE_KEY, API_BASE_DEFAULT).catch(() => {});
      return;
    }
    AsyncStorage.setItem(API_BASE_STORAGE_KEY, normalizedBase).catch(() => {});
  }, [apiBase, apiReady]);

  useEffect(() => {
    let cancelled = false;

    async function loadQueuedUploads() {
      try {
        const raw = await AsyncStorage.getItem(PHOTO_UPLOAD_QUEUE_STORAGE_KEY);
        const parsed = raw ? JSON.parse(raw) : [];
        const normalized = Array.isArray(parsed)
          ? parsed.filter(
              (item) =>
                item &&
                cleanString(item.id) &&
                cleanString(item.claimKey) &&
                cleanString(item.uri)
            )
          : [];
        queuedPhotoUploadsRef.current = normalized;
        if (!cancelled) {
          startTransition(() => {
            setQueuedPhotoUploads(normalized);
          });
        }
      } catch {
        queuedPhotoUploadsRef.current = [];
        if (!cancelled) {
          startTransition(() => {
            setQueuedPhotoUploads([]);
          });
        }
      } finally {
        if (!cancelled) {
          setQueuedPhotoUploadsReady(true);
        }
      }
    }

    loadQueuedUploads();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function loadRoutePlanKeys() {
      try {
        const raw = await AsyncStorage.getItem(ROUTE_PLAN_STORAGE_KEY);
        const parsed = JSON.parse(raw || "[]");
        if (!cancelled && Array.isArray(parsed)) {
          setRoutePlanKeys(parsed.filter((key) => typeof key === "string" && key));
        }
      } catch {}
      if (!cancelled) {
        setRoutePlanReady(true);
      }
    }

    loadRoutePlanKeys();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!routePlanReady) {
      return;
    }
    AsyncStorage.setItem(ROUTE_PLAN_STORAGE_KEY, JSON.stringify(routePlanKeys)).catch(() => {});
  }, [routePlanKeys, routePlanReady]);

  useEffect(() => {
    let cancelled = false;

    async function loadRouteAddressOverrides() {
      try {
        const raw = await AsyncStorage.getItem(ROUTE_ADDRESS_OVERRIDES_STORAGE_KEY);
        const parsed = JSON.parse(raw || "{}");
        if (!cancelled && parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
          setRouteAddressOverrides(parsed);
        }
      } catch {}
      if (!cancelled) {
        setRouteAddressOverridesReady(true);
      }
    }

    loadRouteAddressOverrides();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!routeAddressOverridesReady) {
      return;
    }
    AsyncStorage.setItem(
      ROUTE_ADDRESS_OVERRIDES_STORAGE_KEY,
      JSON.stringify(routeAddressOverrides)
    ).catch(() => {});
  }, [routeAddressOverrides, routeAddressOverridesReady]);

  useEffect(() => {
    if (!apiReady) {
      return;
    }
    fetchJson("/api/mobile/dashboard")
      .then((payload) => {
        startTransition(() => {
          setDashboard(payload);
        });
      })
      .catch(handleError);
  }, [apiBase, apiReady]);

  useEffect(() => {
    if (!apiReady) {
      return;
    }
    loadActiveTab();
  }, [apiBase, homeTab, claimTab, routeStatus, libraryTab, deferredSearch, apiReady]);

  useEffect(() => {
    if (!apiReady) {
      return;
    }
    if (selectedClaimKey) {
      loadClaimDetail(selectedClaimKey);
    } else {
      setClaimDetail(null);
    }
  }, [apiBase, selectedClaimKey, apiReady]);

  useEffect(() => {
    if (!apiReady || !queuedPhotoUploadsReady || !queuedPhotoUploads.length) {
      return;
    }
    const initialAttempt = setTimeout(() => {
      flushQueuedPhotoUploads({ silent: true }).catch(() => {});
    }, 1500);
    const retryTimer = setInterval(() => {
      flushQueuedPhotoUploads({ silent: true }).catch(() => {});
    }, PHOTO_UPLOAD_QUEUE_RETRY_MS);
    return () => {
      clearTimeout(initialAttempt);
      clearInterval(retryTimer);
    };
  }, [apiReady, apiBase, queuedPhotoUploads.length, queuedPhotoUploadsReady]);

  async function persistQueuedPhotoUploads(nextQueue) {
    queuedPhotoUploadsRef.current = nextQueue;
    startTransition(() => {
      setQueuedPhotoUploads(nextQueue);
    });
    try {
      if (nextQueue.length) {
        await AsyncStorage.setItem(PHOTO_UPLOAD_QUEUE_STORAGE_KEY, JSON.stringify(nextQueue));
      } else {
        await AsyncStorage.removeItem(PHOTO_UPLOAD_QUEUE_STORAGE_KEY);
      }
    } catch {}
  }

  async function ensurePhotoQueueDirectory() {
    if (!PHOTO_UPLOAD_QUEUE_DIR) {
      throw new Error("This phone could not prepare local offline photo storage.");
    }
    await FileSystem.makeDirectoryAsync(PHOTO_UPLOAD_QUEUE_DIR, { intermediates: true }).catch(() => {});
    return PHOTO_UPLOAD_QUEUE_DIR;
  }

  async function stagePhotoForOfflineUpload(asset, { claimKey, label = "" }) {
    const uploadAsset = await prepareUploadPhoto(asset);
    const queueDir = await ensurePhotoQueueDirectory();
    const queueId = createPhotoQueueId();
    const safeName = sanitizePhotoQueueName(uploadAsset.name || label || "claim-photo");
    const localUri = `${queueDir}${queueId}-${safeName}.jpg`;
    await FileSystem.copyAsync({ from: uploadAsset.uri, to: localUri });
    const queueItem = {
      id: queueId,
      claimKey,
      label: cleanString(label),
      uri: localUri,
      name: uploadAsset.name || `${safeName}.jpg`,
      type: uploadAsset.type || "image/jpeg",
      capturedAt: uploadAsset.capturedAt || new Date().toISOString(),
      queuedAt: new Date().toISOString(),
      attempts: 0,
      lastError: "",
    };
    await persistQueuedPhotoUploads([queueItem, ...queuedPhotoUploadsRef.current]);
    return queueItem;
  }

  async function flushQueuedPhotoUploads({ specificIds = null, silent = false } = {}) {
    if (queueFlushInFlightRef.current) {
      return { uploaded: [], remaining: queuedPhotoUploadsRef.current };
    }
    const currentQueue = queuedPhotoUploadsRef.current;
    if (!currentQueue.length) {
      return { uploaded: [], remaining: currentQueue };
    }

    const targetIds = specificIds?.length ? new Set(specificIds) : null;
    queueFlushInFlightRef.current = true;
    setQueueFlushing(true);

    const uploaded = [];
    const nextQueue = [];
    const touchedClaimKeys = new Set();

    try {
      for (const item of currentQueue) {
        if (targetIds && !targetIds.has(item.id)) {
          nextQueue.push(item);
          continue;
        }

        try {
          const fileInfo = await FileSystem.getInfoAsync(item.uri);
          if (!fileInfo?.exists) {
            throw new Error("Queued photo file is missing on this phone.");
          }

          const formData = new FormData();
          formData.append("photo", {
            uri: item.uri,
            name: item.name,
            type: item.type || "image/jpeg",
          });

          const labelParam = cleanString(item.label)
            ? `&label=${encodeURIComponent(item.label)}`
            : "";
          const payload = await fetchJson(
            `/api/mobile/upload-photo?key=${encodeURIComponent(item.claimKey)}${labelParam}&captured_at=${encodeURIComponent(item.capturedAt || new Date().toISOString())}`,
            {
              method: "POST",
              body: formData,
            }
          );

          uploaded.push({
            id: item.id,
            claimKey: item.claimKey,
            label: item.label,
            savedName: payload?.file?.name || queueUploadLabel(item),
          });
          touchedClaimKeys.add(item.claimKey);
          await FileSystem.deleteAsync(item.uri, { idempotent: true }).catch(() => {});
        } catch (err) {
          nextQueue.push({
            ...item,
            attempts: Number(item.attempts || 0) + 1,
            lastError: err?.message || "Upload failed.",
            lastAttemptAt: new Date().toISOString(),
          });
        }
      }

      await persistQueuedPhotoUploads(nextQueue);

      if (uploaded.length && touchedClaimKeys.has(selectedClaimKey)) {
        await loadClaimDetail(selectedClaimKey);
      }

      if (!silent && uploaded.length) {
        setUpdateMessage(
          uploaded.length === 1
            ? `${uploaded[0].savedName} uploaded.`
            : `${uploaded.length} queued photos uploaded.`
        );
      }
      return { uploaded, remaining: nextQueue };
    } finally {
      queueFlushInFlightRef.current = false;
      setQueueFlushing(false);
    }
  }

  async function fetchJson(path, options) {
    const primaryBase = apiBase || API_BASE_DEFAULT;
    const tryBases = [primaryBase];
    if (API_BASE_DEFAULT && API_BASE_DEFAULT !== primaryBase) {
      tryBases.push(API_BASE_DEFAULT);
    }
    for (const fallbackBase of API_BASE_FALLBACKS) {
      if (fallbackBase && !tryBases.includes(fallbackBase)) {
        tryBases.push(fallbackBase);
      }
    }

    let lastError = null;
    for (const base of tryBases) {
      try {
        const response = await fetch(`${base}${path}`, options);
        if (!response.ok) {
          let message = `Request failed (${response.status}).`;
          try {
            const payload = await response.json();
            if (payload?.error) {
              message = payload.error;
            }
          } catch {}
          const canTryNext =
            base !== tryBases[tryBases.length - 1] &&
            (response.status >= 500 || (response.status === 404 && isLegacyOrDisallowedPublicBase(base)));
          if (canTryNext) {
            lastError = new Error(message);
            continue;
          }
          throw new Error(message);
        }
        if (base !== apiBase) {
          setApiBase(base);
          AsyncStorage.setItem(API_BASE_STORAGE_KEY, normalizeBaseUrl(base)).catch(() => {});
        }
        setError("");
        return response.json();
      } catch (err) {
        lastError = err;
        if (base === tryBases[tryBases.length - 1]) {
          break;
        }
      }
    }
    throw lastError || new Error("Something went wrong.");
  }

  async function openPhoneAction(phone, mode) {
    const normalizedPhone = digitsOnlyPhone(phone);
    if (!normalizedPhone) {
      setError("No phone number is available for this claim.");
      return;
    }
    try {
      await Linking.openURL(`${mode}:${normalizedPhone}`);
    } catch {
      setError(`${mode === "sms" ? "Text" : "Call"} could not be opened on this phone.`);
    }
  }

  function promptPhoneActions(phone, name) {
    const formattedPhone = formatPhone(phone);
    if (!formattedPhone) {
      Alert.alert(name || "Customer Contact", "No phone number listed.");
      return;
    }
    Alert.alert(name || "Customer Contact", formattedPhone, [
      { text: "Call", onPress: () => openPhoneAction(phone, "tel") },
      { text: "Text", onPress: () => openPhoneAction(phone, "sms") },
      { text: "Cancel", style: "cancel" },
    ]);
  }

  async function promptPhoneActionsForRecord(item, name) {
    try {
      const phone = await ensureClaimPhone(item?.key);
      promptPhoneActions(phone, name);
    } catch (err) {
      handleError(err);
    }
  }

  function handleError(err) {
    setError(err?.message || "Something went wrong.");
  }

  async function checkForAppUpdates() {
    setUpdateChecking(true);
    setUpdateMessage("");
    try {
      if (__DEV__) {
        setUpdateMessage("Update checks work from an installed build or dev client, not inside Expo Go.");
        return;
      }
      if (!Updates.isEnabled) {
        setUpdateMessage("App updates are not configured yet. We still need to connect this app to EAS Update.");
        return;
      }
      const result = await Updates.checkForUpdateAsync();
      if (!result.isAvailable) {
        setUpdateMessage(`No update is available right now. ${currentVersionLabel}`);
        return;
      }
      await Updates.fetchUpdateAsync();
      setUpdateMessage("Update downloaded. Reloading now...");
      await Updates.reloadAsync();
    } catch (err) {
      setUpdateMessage(err?.message || "Could not check for updates.");
    } finally {
      setUpdateChecking(false);
    }
  }

  async function loadActiveTab() {
    const requestId = ++loadSequence.current;
    setLoading(true);
    setError("");
    try {
      if (homeTab === "claims") {
        if (claimTab === "new") {
          startTransition(() => {
            setClaimRecords([]);
          });
          setHasLoadedOnce(true);
          return;
        }
        const params = new URLSearchParams({ status: claimTab, search: deferredSearch });
        const payload = await fetchJson(`/api/mobile/claims?${params.toString()}`);
        if (requestId !== loadSequence.current) {
          return;
        }
        const records = payload.records || [];
        startTransition(() => {
          setClaimRecords(records);
          setDashboard((current) => ({
            ...current,
            summary: payload.summary || current.summary,
          }));
          const firstKey = records[0]?.key || "";
          setSelectedClaimKey((current) =>
            records.some((row) => row.key === current) ? current : firstKey
          );
        });
      } else if (homeTab === "route") {
        const params = new URLSearchParams({ status: routeStatus, search: deferredSearch });
        const routePayload = await fetchJson(`/api/mobile/route-planner?${params.toString()}`);
        if (requestId !== loadSequence.current) {
          return;
        }
        startTransition(() => {
          setRouteData(routePayload);
        });
      } else if (libraryTab === "tools") {
        const payload = await fetchJson("/api/mobile/claim-tools");
        if (requestId !== loadSequence.current) {
          return;
        }
        startTransition(() => {
          setToolsData(payload);
        });
      } else if (libraryTab === "shops") {
        const params = new URLSearchParams({ search: deferredSearch });
        const payload = await fetchJson(`/api/mobile/body-shops?${params.toString()}`);
        if (requestId !== loadSequence.current) {
          return;
        }
        startTransition(() => {
          setBodyShops(payload.records || []);
        });
      } else {
        const params = new URLSearchParams({ search: deferredSearch });
        const payload = await fetchJson(`/api/mobile/insurance-companies?${params.toString()}`);
        if (requestId !== loadSequence.current) {
          return;
        }
        startTransition(() => {
          setInsuranceCompanies(payload.records || []);
        });
      }
      setHasLoadedOnce(true);
    } catch (err) {
      if (requestId === loadSequence.current) {
        handleError(err);
      }
    } finally {
      if (requestId === loadSequence.current) {
        setLoading(false);
      }
    }
  }

  async function loadClaimDetail(key) {
    setDetailLoading(true);
    try {
      const payload = await fetchJson(`/api/mobile/claim?key=${encodeURIComponent(key)}`);
      setClaimDetail(payload);
      if (payload?.key) {
        setClaimPhoneMap((current) =>
          current[payload.key] === (payload.contact_phone || "")
            ? current
            : { ...current, [payload.key]: payload.contact_phone || "" }
        );
      }
      setNoteText("");
      setSelectedPdfPaths([]);
    } catch (err) {
      handleError(err);
    } finally {
      setDetailLoading(false);
    }
  }

  async function ensureClaimPhone(key) {
    if (!key) {
      return "";
    }
    if (Object.prototype.hasOwnProperty.call(claimPhoneMap, key)) {
      return claimPhoneMap[key] || "";
    }
    const payload = await fetchJson(`/api/mobile/claim?key=${encodeURIComponent(key)}`);
    const nextPhone = payload?.contact_phone || "";
    if (payload?.key) {
      setClaimPhoneMap((current) =>
        current[payload.key] === nextPhone ? current : { ...current, [payload.key]: nextPhone }
      );
    }
    return nextPhone;
  }

  function resolvedPhoneForRecord(item) {
    return claimPhoneMap[item?.key] || item?.contact_phone || "";
  }

  function isEmailableFile(file) {
    const lowerName = (file?.name || "").toLowerCase();
    return (
      lowerName.endsWith(".pdf") ||
      lowerName.endsWith(".jpg") ||
      lowerName.endsWith(".jpeg") ||
      lowerName.endsWith(".png") ||
      lowerName.endsWith(".heic") ||
      lowerName.endsWith(".webp")
    );
  }

  function togglePdfSelection(relativePath) {
    setSelectedPdfPaths((current) =>
      current.includes(relativePath)
        ? current.filter((item) => item !== relativePath)
        : [...current, relativePath]
    );
  }

  function openEmailSelectedFiles() {
    if (!claimDetail) {
      return;
    }
    const defaultSubject = [claimDetail.claim_id, claimDetail.customer_name || claimDetail.title, "Claim Files"]
      .filter(Boolean)
      .join(" - ");
    setEmailSubject(defaultSubject);
    setEmailBody("Attached are the selected claim files.");
    setEmailRecipientPickerOpen(false);
    setEmailModalVisible(true);
  }

  function closeEmailModal() {
    if (emailSending) {
      return;
    }
    setEmailRecipientPickerOpen(false);
    setEmailModalVisible(false);
  }

  function mergeEmailAddressList(currentValue, nextEmail) {
    const normalizedNext = cleanString(nextEmail).toLowerCase();
    if (!normalizedNext) {
      return cleanString(currentValue);
    }
    const currentEntries = String(currentValue || "")
      .split(",")
      .map((entry) => cleanString(entry))
      .filter(Boolean);
    const seen = new Set(currentEntries.map((entry) => entry.toLowerCase()));
    if (!seen.has(normalizedNext)) {
      currentEntries.push(nextEmail);
    }
    return currentEntries.join(", ");
  }

  function applyCommonRecipient(recipient, destination = "to") {
    if (!recipient?.email) {
      return;
    }
    if (destination === "cc") {
      setEmailCc((current) => mergeEmailAddressList(current, recipient.email));
      return;
    }
    setEmailTo(recipient.email);
    setEmailRecipientPickerOpen(false);
  }

  async function sendSelectedFilesEmail() {
    if (!selectedClaimKey || !selectedPdfPaths.length) {
      setError("Select at least one file first.");
      return;
    }
    if (!emailTo.trim()) {
      setError("Enter a recipient email first.");
      return;
    }
    setEmailSending(true);
    setError("");
    try {
      const payload = await fetchJson("/api/mobile/send-claim-files-email", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          key: selectedClaimKey,
          to: emailTo.trim(),
          cc: emailCc.trim(),
          subject: emailSubject.trim(),
          body: emailBody,
          paths: selectedPdfPaths,
        }),
      });
      setUpdateMessage(
        payload?.attachment_count
          ? `${payload.attachment_count} file${payload.attachment_count === 1 ? "" : "s"} emailed to ${payload.sent_to}.`
          : "Email sent."
      );
      setEmailModalVisible(false);
      setSelectedPdfPaths([]);
    } catch (err) {
      handleError(err);
    } finally {
      setEmailSending(false);
    }
  }

  async function saveNote() {
    if (!selectedClaimKey || !noteText.trim()) {
      return;
    }
    setNoteSaving(true);
    try {
      const payload = await fetchJson("/api/mobile/claim-note", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ key: selectedClaimKey, text: noteText.trim() }),
      });
      if (payload.claim) {
        setClaimDetail(payload.claim);
      }
      setNoteText("");
    } catch (err) {
      handleError(err);
    } finally {
      setNoteSaving(false);
    }
  }

  async function uploadNewAssignmentPdf() {
    if (assignmentUploadInFlight) {
      return;
    }
    setError("");
    setUpdateMessage("");
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: "application/pdf",
        copyToCacheDirectory: true,
        multiple: false,
      });
      if (result.canceled || !result.assets?.length) {
        return;
      }
      const asset = result.assets[0];
      if (!asset?.uri) {
        throw new Error("No assignment PDF was selected.");
      }
      setAssignmentUploadInFlight(true);
      setAssignmentUploadResult(null);
      const formData = new FormData();
      formData.append("assignment", {
        uri: asset.uri,
        name: asset.name || "assign.pdf",
        type: asset.mimeType || "application/pdf",
      });
      const payload = await fetchJson("/api/mobile/upload-assignment-pdf", {
        method: "POST",
        body: formData,
      });
      setAssignmentUploadResult(payload);
      const importedClaim = payload?.claim || null;
      const importedLabel =
        [importedClaim?.claim_id, importedClaim?.customer_name || importedClaim?.title]
          .filter(Boolean)
          .join(" - ") ||
        payload?.folder_name ||
        "Pending claim";
      setUpdateMessage(`Assignment imported into ${importedLabel}.`);
      if (importedClaim?.claim_id) {
        setSearch(importedClaim.claim_id);
      }
      setClaimTab("open");
      if (importedClaim?.key) {
        setSelectedClaimKey(importedClaim.key);
        await loadClaimDetail(importedClaim.key);
        setDetailVisible(true);
      }
    } catch (err) {
      handleError(err);
    } finally {
      setAssignmentUploadInFlight(false);
    }
  }

  function resetPhotoSessionState() {
    setPendingPhotoAsset(null);
    setPendingPhotoLabel("");
    setPhotoSessionUploads([]);
    setPhotoZoom(0);
    setCameraReady(false);
    setPhotoCaptureCueVisible(false);
    setPhotoCaptureCueText("");
  }

  function showPhotoCaptureCue(label) {
    if (photoCaptureCueTimeoutRef.current) {
      clearTimeout(photoCaptureCueTimeoutRef.current);
    }
    Vibration.vibrate(90);
    setPhotoCaptureCueText(label ? `${label} captured` : "Photo captured");
    setPhotoCaptureCueVisible(true);
    photoCaptureCueTimeoutRef.current = setTimeout(() => {
      setPhotoCaptureCueVisible(false);
      setPhotoCaptureCueText("");
      photoCaptureCueTimeoutRef.current = null;
    }, 1100);
  }

  async function closePhotoSession() {
    setPhotoSessionVisible(false);
    resetPhotoSessionState();
    if (selectedClaimKey) {
      await loadClaimDetail(selectedClaimKey);
    }
  }

  function startLabeledPhotoCapture(label) {
    setPendingPhotoLabel(label);
    setPendingPhotoAsset(null);
    setCameraReady(false);
    setError("");
  }

  function returnToPhotoLabelPicker() {
    setPendingPhotoLabel("");
    setPendingPhotoAsset(null);
    setCameraReady(false);
  }

  async function capturePhotoForSession() {
    if (!selectedClaimKey) {
      return;
    }
    if (!pendingPhotoLabel) {
      setError("Pick a label first.");
      return;
    }
    if (!cameraPermission?.granted) {
      const permission = await requestCameraPermission();
      if (!permission.granted) {
        setError("Camera permission is required.");
        return;
      }
    }
    if (!cameraReady || !cameraRef.current?.takePictureAsync) {
      setError("Camera is still getting ready.");
      return;
    }
    setUploadingPhoto(true);
    try {
      const result = await cameraRef.current.takePictureAsync({
        quality: 0.9,
        shutterSound: false,
      });
      if (!result?.uri) {
        return;
      }
      showPhotoCaptureCue(pendingPhotoLabel);
      setError("");
      setUpdateMessage(`${pendingPhotoLabel} captured. Saving now...`);
      const queuedItem = await stagePhotoForOfflineUpload(result, {
        claimKey: selectedClaimKey,
        label: pendingPhotoLabel,
      });
      const flushResult = await flushQueuedPhotoUploads({
        specificIds: [queuedItem.id],
        silent: true,
      });
      const uploadedItem = flushResult.uploaded.find((item) => item.id === queuedItem.id);
      const savedName = uploadedItem?.savedName || queueUploadLabel(queuedItem);
      startTransition(() => {
        setPhotoSessionUploads((current) => [
          uploadedItem ? savedName : `${savedName} (queued)`,
          ...current,
        ].slice(0, 12));
      });
      setUpdateMessage(
        uploadedItem
          ? `${savedName} saved.`
          : `${savedName} queued on this phone until your reception is good enough to upload it.`
      );
      returnToPhotoLabelPicker();
    } catch (err) {
      handleError(err);
    } finally {
      setUploadingPhoto(false);
    }
  }

  async function openPhotoSession() {
    if (!selectedClaimKey) {
      return;
    }
    const permission = cameraPermission?.granted ? cameraPermission : await requestCameraPermission();
    if (!permission?.granted) {
      setError("Camera permission is required.");
      return;
    }
    resetPhotoSessionState();
    setPhotoSessionVisible(true);
    setError("");
  }

  async function savePendingPhoto({ finishAfter = false } = {}) {
    if (!selectedClaimKey || !pendingPhotoAsset) {
      return;
    }
    if (!pendingPhotoLabel) {
      setError("Pick a photo label before saving.");
      return;
    }

    try {
      setUploadingPhoto(true);
      setError("");
      const queuedItem = await stagePhotoForOfflineUpload(pendingPhotoAsset, {
        claimKey: selectedClaimKey,
        label: pendingPhotoLabel,
      });
      const flushResult = await flushQueuedPhotoUploads({
        specificIds: [queuedItem.id],
        silent: true,
      });
      const uploadedItem = flushResult.uploaded.find((item) => item.id === queuedItem.id);
      const savedName = uploadedItem?.savedName || queueUploadLabel(queuedItem);
      startTransition(() => {
        setPhotoSessionUploads((current) => [
          uploadedItem ? savedName : `${savedName} (queued)`,
          ...current,
        ].slice(0, 12));
      });
      setCameraReady(false);
      setPendingPhotoAsset(null);
      setPendingPhotoLabel("");
      setUpdateMessage(
        uploadedItem
          ? `${savedName} saved.`
          : `${savedName} queued on this phone until your reception is good enough to upload it.`
      );

      if (finishAfter) {
        await closePhotoSession();
      }
    } catch (err) {
      handleError(err);
    } finally {
      setUploadingPhoto(false);
    }
  }

  async function retakePendingPhoto() {
    setPendingPhotoAsset(null);
    setPendingPhotoLabel("");
    setCameraReady(false);
  }

  async function deleteClaimFile(file) {
    if (!selectedClaimKey || !file?.relative_path) {
      return;
    }
    Alert.alert(
      "Delete File",
      `Delete ${file.name}?`,
      [
        { text: "Cancel", style: "cancel" },
        {
          text: "Delete",
          style: "destructive",
          onPress: async () => {
            try {
              setError("");
              const payload = await fetchJson("/api/mobile/delete-file", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ key: selectedClaimKey, path: file.relative_path }),
              });
              if (payload.claim) {
                setClaimDetail(payload.claim);
              } else {
                await loadClaimDetail(selectedClaimKey);
              }
              startTransition(() => {
                setSelectedPdfPaths((current) => current.filter((path) => path !== file.relative_path));
              });
              setUpdateMessage(`${file.name} deleted.`);
            } catch (err) {
              handleError(err);
            }
          },
        },
      ]
    );
  }

  async function uploadPhotosFromLibrary() {
    if (!selectedClaimKey) {
      return;
    }
    setUploadingPhoto(true);
    try {
      const permission = await ImagePicker.requestMediaLibraryPermissionsAsync();
      if (!permission.granted) {
        throw new Error("Photo library permission is required.");
      }

      const result = await ImagePicker.launchImageLibraryAsync({
        quality: 0.9,
        allowsMultipleSelection: true,
        selectionLimit: 0,
      });

      if (result.canceled || !result.assets?.length) {
        return;
      }

      const stagedItems = [];
      for (const asset of result.assets) {
        const queuedItem = await stagePhotoForOfflineUpload(asset, {
          claimKey: selectedClaimKey,
        });
        stagedItems.push(queuedItem);
      }

      const flushResult = await flushQueuedPhotoUploads({
        specificIds: stagedItems.map((item) => item.id),
        silent: true,
      });
      const uploadedCount = flushResult.uploaded.length;
      const queuedCount = stagedItems.length - uploadedCount;

      if (uploadedCount && queuedCount) {
        setUpdateMessage(`${uploadedCount} photo${uploadedCount === 1 ? "" : "s"} uploaded, ${queuedCount} queued until signal improves.`);
      } else if (uploadedCount) {
        setUpdateMessage(
          uploadedCount > 1 ? `${uploadedCount} photos uploaded.` : "Photo uploaded."
        );
      } else {
        setUpdateMessage(
          stagedItems.length > 1
            ? `${stagedItems.length} photos queued on this phone until your reception is good enough to upload them.`
            : "Photo queued on this phone until your reception is good enough to upload it."
        );
      }
    } catch (err) {
      handleError(err);
    } finally {
      setUploadingPhoto(false);
    }
  }

  async function retryCurrentClaimQueuedUploads() {
    if (!currentClaimQueuedUploads.length) {
      return;
    }
    try {
      setError("");
      const result = await flushQueuedPhotoUploads({
        specificIds: currentClaimQueuedUploads.map((item) => item.id),
      });
      if (!result.uploaded.length) {
        setUpdateMessage("Still waiting on better reception to upload the queued photos.");
      }
    } catch (err) {
      handleError(err);
    }
  }

  async function retryAllQueuedUploads() {
    if (!queuedPhotoUploads.length) {
      return;
    }
    try {
      setError("");
      const result = await flushQueuedPhotoUploads();
      if (!result.uploaded.length) {
        setUpdateMessage("Still waiting on better reception to upload the queued photos.");
      }
    } catch (err) {
      handleError(err);
    }
  }

  async function saveRoutePlan(keys, successMessage) {
    setRouteSaving(true);
    setError("");
    try {
      startTransition(() => {
        setRoutePlanKeys([...new Set(keys.filter(Boolean))]);
      });
      if (successMessage) {
        setUpdateMessage(successMessage);
      }
    } catch (err) {
      handleError(err);
    } finally {
      setRouteSaving(false);
    }
  }

  async function addRouteStop(item) {
    const nextKeys = [...routePlanKeys, item.key];
    await saveRoutePlan(nextKeys, "Stop added to route.");
  }

  async function removeRouteStop(key) {
    const nextKeys = routePlanKeys.filter((recordKey) => recordKey !== key);
    await saveRoutePlan(nextKeys, "Stop removed from route.");
  }

  async function moveRouteStop(key, direction) {
    const index = routePlanKeys.indexOf(key);
    const swapIndex = index + direction;
    if (index < 0 || swapIndex < 0 || swapIndex >= routePlanKeys.length) {
      return;
    }
    const nextKeys = [...routePlanKeys];
    [nextKeys[index], nextKeys[swapIndex]] = [nextKeys[swapIndex], nextKeys[index]];
    await saveRoutePlan(nextKeys, "Route order updated.");
  }

  async function clearRoutePlan() {
    await saveRoutePlan([], "Route cleared.");
  }

  function openRouteClaim(key) {
    setSelectedClaimKey(key);
    setDetailVisible(true);
  }

  function currentRouteOrigin() {
    return (routeData.start_from || dashboard.settings?.route_home_address || "").trim();
  }

  function routeStopAddresses() {
    return plannedRouteRecords.map((record) => record.address?.trim()).filter(Boolean);
  }

  async function openGoogleMapsRoute() {
    const origin = currentRouteOrigin();
    const stops = routeStopAddresses();

    if (!origin) {
      setError("Set a route start address in the desktop app first.");
      return;
    }
    if (!stops.length) {
      setError("Add at least one claim to the route first.");
      return;
    }

    const routeParts = [origin, ...stops].map((part) => encodeURIComponent(part));
    const routeUrl = `https://www.google.com/maps/dir/${routeParts.join("/")}/`;
    try {
      await Linking.openURL(routeUrl);
    } catch {
      setError("Google Maps could not be opened on this phone.");
    }
  }

  async function openWazeRoute() {
    const destination = plannedRouteRecords[plannedRouteRecords.length - 1]?.address?.trim();
    if (!destination) {
      setError("Add at least one claim with an address first.");
      return;
    }

    const routeUrl = `https://waze.com/ul?q=${encodeURIComponent(destination)}&navigate=yes`;
    try {
      await Linking.openURL(routeUrl);
    } catch {
      setError("Waze could not be opened on this phone.");
    }
  }

  async function openAppleMapsRoute() {
    const origin = currentRouteOrigin();
    const destination = plannedRouteRecords[plannedRouteRecords.length - 1]?.address?.trim();
    if (!origin) {
      setError("Set a route start address in the desktop app first.");
      return;
    }
    if (!destination) {
      setError("Add at least one claim with an address first.");
      return;
    }

    const routeUrl =
      `http://maps.apple.com/?saddr=${encodeURIComponent(origin)}` +
      `&daddr=${encodeURIComponent(destination)}&dirflg=d`;
    try {
      await Linking.openURL(routeUrl);
    } catch {
      setError("Apple Maps could not be opened on this phone.");
    }
  }

  async function openClaimLocation(address) {
    const normalizedAddress = cleanString(address);
    if (!normalizedAddress || normalizedAddress === "-") {
      setError("No vehicle address is listed on this claim yet.");
      return;
    }
    const geoUrl = `geo:0,0?q=${encodeURIComponent(normalizedAddress)}`;
    const appleUrl = `http://maps.apple.com/?q=${encodeURIComponent(normalizedAddress)}`;
    const googleUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(normalizedAddress)}`;
    try {
      if (Platform.OS === "android" && (await Linking.canOpenURL(geoUrl))) {
        await Linking.openURL(geoUrl);
        return;
      }
      if (Platform.OS === "ios" && (await Linking.canOpenURL(appleUrl))) {
        await Linking.openURL(appleUrl);
        return;
      }
      await Linking.openURL(googleUrl);
    } catch {
      setError("Could not open maps for this claim.");
    }
  }

  function startEditingRouteAddress(item) {
    setEditingRouteKey(item.key);
    setEditingRouteAddress(item.address || "");
  }

  function cancelEditingRouteAddress() {
    setEditingRouteKey("");
    setEditingRouteAddress("");
  }

  async function saveEditedRouteAddress() {
    if (!editingRouteKey) {
      return;
    }
    setRouteAddressSaving(true);
    try {
      const nextAddress = editingRouteAddress.trim();
      startTransition(() => {
        setRouteAddressOverrides((current) => {
          const next = { ...current };
          if (nextAddress) {
            next[editingRouteKey] = nextAddress;
          } else {
            delete next[editingRouteKey];
          }
          return next;
        });
      });
      setUpdateMessage("Route address updated.");
      cancelEditingRouteAddress();
    } catch (err) {
      handleError(err);
    } finally {
      setRouteAddressSaving(false);
    }
  }

  const detailRows = useMemo(() => {
    if (!claimDetail) {
      return [];
    }
    return [
      { label: "Claim ID", value: claimDetail.claim_id },
      { label: "Customer", value: claimDetail.customer_name || claimDetail.title },
      {
        label: "Phone",
        value: formatPhone(claimDetail.contact_phone) || "No phone number listed",
        phone: claimDetail.contact_phone,
      },
      { label: "Insurance", value: claimDetail.insurance_company },
      { label: "Claim #", value: claimDetail.claim_number },
      { label: "Status", value: claimDetail.status },
      { label: "Type", value: claimDetail.claim_type },
      { label: "Date Of Loss", value: claimDetail.date_of_loss },
      { label: "Town", value: claimDetail.town },
      { label: "Shop", value: claimDetail.shop_name },
      { label: "Vehicle", value: claimDetail.vehicle },
      { label: "VIN", value: claimDetail.vin },
      { label: "Inspection Location", value: claimDetail.route_display_address },
    ];
  }, [claimDetail]);

  const routeRecords = useMemo(() => {
    const recordMap = new Map();
    [...(routeData.records || []), ...(routeData.planned_records || [])].forEach((record) => {
      if (record?.key) {
        recordMap.set(record.key, {
          ...record,
          address: routeAddressOverrides[record.key]?.trim() || record.address,
        });
      }
    });
    return Array.from(recordMap.values());
  }, [routeAddressOverrides, routeData]);

  useEffect(() => {
    if (!routeRecords.length) {
      return;
    }
    const availableKeys = new Set(routeRecords.map((record) => record.key));
    setRoutePlanKeys((current) => {
      const next = current.filter((key) => availableKeys.has(key));
      return next.length === current.length ? current : next;
    });
  }, [routeRecords]);

  const plannedRouteRecords = useMemo(() => {
    const recordMap = new Map(routeRecords.map((record) => [record.key, record]));
    return routePlanKeys.map((key) => recordMap.get(key)).filter(Boolean);
  }, [routePlanKeys, routeRecords]);

  useEffect(() => {
    if (!apiReady) {
      return;
    }
    const targets = [...claimRecords, ...plannedRouteRecords]
      .filter((item) => item?.key && !resolvedPhoneForRecord(item))
      .slice(0, 40);
    if (!targets.length) {
      return;
    }
    let cancelled = false;
    (async () => {
      for (const item of targets) {
        if (cancelled) {
          return;
        }
        try {
          await ensureClaimPhone(item.key);
        } catch {
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [apiReady, claimRecords, plannedRouteRecords, claimPhoneMap]);

  const availableRouteRecords = useMemo(() => {
    const selectedKeys = new Set(routePlanKeys);
    const normalizedSearch = deferredSearch.trim().toLowerCase();
    return routeRecords.filter((record) => {
      if (selectedKeys.has(record.key)) {
        return false;
      }
      if (routeStatus !== "all" && (record.status || "").toLowerCase() !== routeStatus) {
        return false;
      }
      if (!normalizedSearch) {
        return true;
      }
      const haystack = [
        record.claim_id,
        record.customer_name,
        record.address,
        record.shop_name,
        record.town,
        record.vehicle,
        record.insurance_company,
        record.status_label,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return haystack.includes(normalizedSearch);
    });
  }, [deferredSearch, routePlanKeys, routeRecords, routeStatus]);

  return (
    <SafeAreaView style={styles.safe}>
      <ExpoStatusBar style="dark" backgroundColor="#eef4fb" translucent={false} />
      <View style={styles.shell}>
        <View style={styles.headerTop}>
          <View style={styles.header}>
            <Text style={styles.eyebrow}>{APP_OWNER_LABEL}</Text>
            <Text style={styles.title}>Claim Manager Companion</Text>
            <Text style={styles.subtitle}>
              Field access for claims, files, notes, routes, and photo upload.
            </Text>
            <Text style={styles.versionText}>{currentVersionLabel}</Text>
            <Text style={styles.versionSubtext}>{currentVersionDetail}</Text>
          </View>
          <Pressable style={styles.ghostBtn} onPress={() => setUtilityOpen(true)}>
            <Text style={styles.ghostBtnText}>Settings</Text>
          </Pressable>
        </View>

        <View style={styles.statusRow}>
          <View style={styles.statusChip}>
            <Text style={styles.statusChipLabel}>PC</Text>
            <Text style={styles.statusChipValue} numberOfLines={1}>
              {apiBase.replace(/^https?:\/\//, "")}
            </Text>
          </View>
          {queuedPhotoUploads.length ? (
            <Pressable style={[styles.statusChip, styles.pendingStatusChip]} onPress={retryAllQueuedUploads}>
              <Text style={styles.statusChipLabel}>Queued</Text>
              <Text style={styles.statusChipValue} numberOfLines={1}>
                {queueFlushing ? "Uploading..." : `${queuedPhotoUploads.length} photo${queuedPhotoUploads.length === 1 ? "" : "s"}`}
              </Text>
            </Pressable>
          ) : null}
          {updateMessage ? (
            <Text style={styles.statusMessage} numberOfLines={1}>
              {updateMessage}
            </Text>
          ) : null}
        </View>

        <View style={styles.summaryRow}>
          <Summary label="All" value={dashboard.summary?.all} />
          <Summary label="Open" value={dashboard.summary?.open} />
          <Summary label="Closed" value={dashboard.summary?.closed} />
        </View>

        <View style={styles.tabRow}>
          {HOME_TABS.map(([key, label]) => (
            <Chip key={key} label={label} active={homeTab === key} onPress={() => setHomeTab(key)} />
          ))}
        </View>

        {!(homeTab === "claims" && claimTab === "new") ? (
          <TextInput
            value={search}
            onChangeText={setSearch}
            autoCapitalize="none"
            autoCorrect={false}
            style={styles.search}
            placeholder={
              homeTab === "library"
                ? "Search library..."
                : homeTab === "route"
                  ? "Search any claim for your route..."
                  : "Search claims, insurance, town, shop..."
            }
          />
        ) : null}

        {homeTab === "claims" ? (
          <>
            <View style={styles.tabRow}>
              {CLAIM_TABS.map(([key, label]) => (
                <Chip key={key} label={label} active={claimTab === key} onPress={() => setClaimTab(key)} />
              ))}
            </View>
            <Text style={styles.inlineHint}>
              {claimTab === "new"
                ? "Pick a new assignment PDF from your phone and send it to the Home PC so it creates the pending claim folder for you."
                : "Tap any claim to open files, photos, notes, and quick actions."}
            </Text>
          </>
        ) : null}

        {homeTab === "library" ? (
          <View style={styles.tabRow}>
            {LIBRARY_TABS.map(([key, label]) => (
              <Chip key={key} label={label} active={libraryTab === key} onPress={() => setLibraryTab(key)} />
            ))}
          </View>
        ) : null}

        {error ? (
          <View style={styles.error}>
            <Text style={styles.errorText}>{error}</Text>
          </View>
        ) : null}

        {!hasLoadedOnce && loading ? (
          <View style={styles.loading}>
            <ActivityIndicator size="large" color="#0e3a66" />
            <Text style={styles.helper}>Loading from your PC...</Text>
          </View>
        ) : (
          <View style={styles.contentWrap}>
            {renderContent()}
            {loading ? (
              <View style={styles.loadingOverlay} pointerEvents="none">
                <View style={styles.loadingPill}>
                  <ActivityIndicator size="small" color="#0e3a66" />
                  <Text style={styles.loadingPillText}>Refreshing...</Text>
                </View>
              </View>
            ) : null}
          </View>
        )}
      </View>

      <Modal visible={utilityOpen} animationType="slide" transparent onRequestClose={() => setUtilityOpen(false)}>
        <View style={styles.sheetBackdrop}>
          <Pressable style={styles.sheetDismiss} onPress={() => setUtilityOpen(false)} />
          <SafeAreaView style={styles.sheetCard}>
            <View style={styles.modalHead}>
              <Pressable style={styles.smallBtn} onPress={() => setUtilityOpen(false)}>
                <Text style={styles.smallBtnText}>Close</Text>
              </Pressable>
              <Text style={styles.modalTitle}>Connection</Text>
              <Pressable style={styles.smallBtn} onPress={checkForAppUpdates}>
                <Text style={styles.smallBtnText}>{updateChecking ? "..." : "Update"}</Text>
              </Pressable>
            </View>

            <View style={styles.utilityCard}>
              <Text style={styles.utilityHint}>{CONNECTION_HINT}</Text>
              <Text style={styles.helperStrong}>{currentVersionLabel}</Text>
              <Text style={styles.helper}>{currentVersionDetail}</Text>
              {CONNECTION_LOCKED ? (
                <View style={styles.input}>
                  <Text style={styles.helperStrong}>{apiBase}</Text>
                </View>
              ) : (
                <TextInput
                  value={apiBase}
                  onChangeText={setApiBase}
                  autoCapitalize="none"
                  autoCorrect={false}
                  style={styles.input}
                  placeholder="https://api.luxuryimportsusa.shop"
                />
              )}
              <Text style={styles.helper}>Public address: {API_BASE_DEFAULT}</Text>
              {API_BASE_FALLBACKS.length ? (
                <Text style={styles.helper}>Fallbacks: {API_BASE_FALLBACKS.join(", ")}</Text>
              ) : null}
              {updateMessage ? <Text style={styles.helper}>{updateMessage}</Text> : null}
            </View>
          </SafeAreaView>
        </View>
      </Modal>

      <Modal visible={detailVisible} animationType="slide" onRequestClose={() => setDetailVisible(false)}>
        <SafeAreaView style={styles.safe}>
          <View style={styles.shell}>
            <View style={styles.modalHead}>
              <Pressable style={styles.smallBtn} onPress={() => setDetailVisible(false)}>
                <Text style={styles.smallBtnText}>Back</Text>
              </Pressable>
              <Text style={styles.modalTitle}>Claim Detail</Text>
              <Pressable style={styles.smallBtn} onPress={() => loadClaimDetail(selectedClaimKey)}>
                <Text style={styles.smallBtnText}>Refresh</Text>
              </Pressable>
            </View>

            {detailLoading || !claimDetail ? (
              <View style={styles.loading}>
                <ActivityIndicator size="large" color="#0e3a66" />
              </View>
            ) : (
              <ScrollView contentContainerStyle={styles.scrollPad}>
                <View style={styles.card}>
                  <Text style={styles.section}>{claimDetail.customer_name || claimDetail.title || "Claim"}</Text>
                  {detailRows.map((row) => (
                    <FieldRow
                      key={row.label}
                      label={row.label}
                      value={row.value}
                      onPress={
                        row.label === "Phone"
                          ? () => promptPhoneActions(row.phone, claimDetail.customer_name || claimDetail.title)
                          : row.label === "Inspection Location"
                            ? () => openClaimLocation(row.value)
                          : undefined
                      }
                      onLongPress={
                        row.label === "Phone"
                          ? () => promptPhoneActions(row.phone, claimDetail.customer_name || claimDetail.title)
                          : row.label === "Inspection Location"
                            ? () => openClaimLocation(row.value)
                          : undefined
                      }
                    />
                  ))}
                </View>

                <View style={styles.card}>
                  <Text style={styles.section}>Actions</Text>
                  <View style={styles.stack}>
                    {claimDetail.assignment_url ? (
                      <Action label="Open Assign Sheet" primary onPress={() => Linking.openURL(claimDetail.assignment_url)} />
                    ) : null}
                  </View>
                </View>

                <View style={styles.card}>
                  <Text style={styles.section}>Photos</Text>
                  <Text style={styles.helper}>
                    Save photos straight into this claim folder on your PC. The photo session keeps you in a continuous capture flow.
                  </Text>
                  {currentClaimQueuedUploads.length ? (
                    <View style={styles.queueNotice}>
                      <Text style={styles.queueNoticeTitle}>
                        {currentClaimQueuedUploads.length} photo{currentClaimQueuedUploads.length === 1 ? "" : "s"} waiting on this phone
                      </Text>
                      <Text style={styles.helper}>
                        Reception was not good enough yet. These photos are stored locally and will upload automatically when the connection is strong enough.
                      </Text>
                      <Action
                        label={queueFlushing ? "Uploading..." : "Retry Pending Uploads"}
                        onPress={retryCurrentClaimQueuedUploads}
                        disabled={queueFlushing}
                      />
                    </View>
                  ) : null}
                  <View style={styles.photoRow}>
                    <Action
                      label={uploadingPhoto ? "Working..." : "Take Photos"}
                      onPress={openPhotoSession}
                      compact
                      disabled={uploadingPhoto}
                    />
                    <Action
                      label={uploadingPhoto ? "Working..." : "Upload Photos"}
                      onPress={uploadPhotosFromLibrary}
                      compact
                      disabled={uploadingPhoto}
                    />
                  </View>
                </View>

                <View style={styles.card}>
                  <Text style={styles.section}>Notes</Text>
                  <Text style={styles.noteBlock}>{claimDetail.notes || "No saved notes yet."}</Text>
                  <TextInput
                    value={noteText}
                    onChangeText={setNoteText}
                    multiline
                    style={styles.notesInput}
                    placeholder="Add a field note..."
                  />
                  <Action label={noteSaving ? "Saving..." : "Save Note"} primary onPress={saveNote} />
                </View>

                <View style={styles.card}>
                  <Text style={styles.section}>Files In Claim Folder</Text>
                  <Text style={styles.helper}>
                    Select PDF or photo files you want to email from this claim folder, then send them from the Home PC.
                  </Text>
                  <View style={styles.stack}>
                    <Action
                      label={emailSending ? "Sending..." : "Email Selected Files"}
                      primary
                      onPress={openEmailSelectedFiles}
                      disabled={!selectedPdfPaths.length || emailSending}
                    />
                  </View>
                  {(claimDetail.files || []).length ? (
                    claimDetail.files.map((file) => (
                      <View key={file.relative_path} style={[styles.listItem, styles.fileListItem]}>
                        <Pressable style={styles.fileMainArea} onPress={() => Linking.openURL(file.file_url)}>
                          <Text style={styles.itemTitle}>{file.name}</Text>
                          <Text style={styles.itemMeta}>{file.relative_path}</Text>
                        </Pressable>
                        <View style={styles.fileActionColumn}>
                          {isEmailableFile(file) ? (
                            <Pressable
                              onPress={() => togglePdfSelection(file.relative_path)}
                              style={[
                                styles.fileSelectChip,
                                selectedPdfPaths.includes(file.relative_path) && styles.fileSelectChipActive,
                              ]}
                            >
                              <Text
                                style={[
                                  styles.fileSelectChipText,
                                  selectedPdfPaths.includes(file.relative_path) && styles.fileSelectChipTextActive,
                                ]}
                              >
                                {selectedPdfPaths.includes(file.relative_path) ? "Selected" : "Select File"}
                              </Text>
                            </Pressable>
                          ) : (
                            <Text style={styles.fileNonPdfHint}>Open only</Text>
                          )}
                          <Pressable style={styles.fileDeleteChip} onPress={() => deleteClaimFile(file)}>
                            <Text style={styles.fileDeleteChipText}>Delete</Text>
                          </Pressable>
                        </View>
                      </View>
                    ))
                  ) : (
                    <Text style={styles.helper}>No files found in this claim folder.</Text>
                  )}
                </View>
              </ScrollView>
            )}
          </View>
        </SafeAreaView>
      </Modal>

      <Modal visible={emailModalVisible} animationType="slide" transparent onRequestClose={closeEmailModal}>
        <View style={styles.sheetBackdrop}>
          <Pressable style={styles.sheetDismiss} onPress={closeEmailModal} />
          <SafeAreaView style={styles.sheetCard}>
            <View style={styles.modalHead}>
              <Pressable style={styles.smallBtn} onPress={closeEmailModal} disabled={emailSending}>
                <Text style={styles.smallBtnText}>Close</Text>
              </Pressable>
              <Text style={styles.modalTitle}>Email Files</Text>
              <Action
                label={emailSending ? "..." : "Send"}
                primary
                compact
                onPress={sendSelectedFilesEmail}
                disabled={emailSending || !selectedPdfPaths.length}
              />
            </View>

            <View style={styles.utilityCard}>
              <Text style={styles.helperStrong}>
                {selectedPdfPaths.length} file{selectedPdfPaths.length === 1 ? "" : "s"} selected
              </Text>
              <TextInput
                value={emailTo}
                onChangeText={setEmailTo}
                autoCapitalize="none"
                autoCorrect={false}
                keyboardType="email-address"
                style={styles.input}
                placeholder="To"
              />
              <View style={styles.recipientPickerWrap}>
                <Pressable
                  style={styles.recipientPickerToggle}
                  onPress={() => setEmailRecipientPickerOpen((current) => !current)}
                >
                  <Text style={styles.recipientPickerToggleText}>
                    {emailRecipientPickerOpen ? "Hide Common Recipients" : "Common Recipients"}
                  </Text>
                </Pressable>
                {emailRecipientPickerOpen ? (
                  <View style={styles.recipientPickerMenu}>
                    <Text style={styles.helper}>Tap a name to fill To, or add them to Cc.</Text>
                    {COMMON_EMAIL_RECIPIENTS.map((recipient) => (
                      <View key={recipient.email} style={styles.recipientPickerRow}>
                        <View style={styles.recipientPickerInfo}>
                          <Text style={styles.recipientPickerName}>{recipient.label}</Text>
                          <Text style={styles.recipientPickerEmail}>{recipient.email}</Text>
                        </View>
                        <View style={styles.recipientPickerActions}>
                          <Pressable
                            style={styles.recipientPickerChip}
                            onPress={() => applyCommonRecipient(recipient, "to")}
                          >
                            <Text style={styles.recipientPickerChipText}>To</Text>
                          </Pressable>
                          <Pressable
                            style={styles.recipientPickerChip}
                            onPress={() => applyCommonRecipient(recipient, "cc")}
                          >
                            <Text style={styles.recipientPickerChipText}>Cc</Text>
                          </Pressable>
                        </View>
                      </View>
                    ))}
                  </View>
                ) : null}
              </View>
              <TextInput
                value={emailCc}
                onChangeText={setEmailCc}
                autoCapitalize="none"
                autoCorrect={false}
                keyboardType="email-address"
                style={styles.input}
                placeholder="Cc (optional)"
              />
              <TextInput
                value={emailSubject}
                onChangeText={setEmailSubject}
                autoCapitalize="sentences"
                autoCorrect={true}
                style={styles.input}
                placeholder="Subject"
              />
              <TextInput
                value={emailBody}
                onChangeText={setEmailBody}
                multiline
                style={styles.notesInput}
                placeholder="Message"
              />
              <View style={styles.emailSelectionList}>
                {selectedPdfPaths.map((path) => (
                  <Text key={path} style={styles.emailSelectionItem}>
                    {path}
                  </Text>
                ))}
              </View>
            </View>
          </SafeAreaView>
        </View>
      </Modal>

      <Modal visible={photoSessionVisible} animationType="slide" onRequestClose={closePhotoSession}>
        <SafeAreaView style={styles.safe}>
          <View style={styles.shell}>
            <View style={styles.modalHead}>
              <Pressable style={styles.smallBtn} onPress={closePhotoSession}>
                <Text style={styles.smallBtnText}>Back</Text>
              </Pressable>
              <Text style={styles.modalTitle}>Photo Session</Text>
              <Pressable style={styles.smallBtn} onPress={closePhotoSession}>
                <Text style={styles.smallBtnText}>Done</Text>
              </Pressable>
            </View>

            <ScrollView contentContainerStyle={styles.scrollPad}>
              <View style={styles.card}>
                <Text style={styles.section}>{claimDetail?.customer_name || claimDetail?.title || "Claim"}</Text>
                <Text style={styles.helper}>
                  The rear camera opens first. After each shot, pick the photo label, save it, and keep going until you are done.
                </Text>
                {photoCaptureCueVisible ? (
                  <View style={styles.captureCueBanner}>
                    <Text style={styles.captureCueText}>{photoCaptureCueText || "Photo captured"}</Text>
                  </View>
                ) : null}
                <Text style={styles.helperStrong}>
                  {photoSessionUploads.length
                    ? `${photoSessionUploads.length} photo${photoSessionUploads.length === 1 ? "" : "s"} saved in this session.`
                    : "No photos saved in this session yet."}
                </Text>
                {currentClaimQueuedUploads.length ? (
                  <Text style={styles.helper}>
                    {currentClaimQueuedUploads.length} photo{currentClaimQueuedUploads.length === 1 ? "" : "s"} currently stored on this phone and waiting for better reception before uploading.
                  </Text>
                ) : null}
                {photoSessionUploads.length ? (
                  <View style={styles.photoSessionSavedList}>
                    {photoSessionUploads.map((name) => (
                      <Text key={name} style={styles.photoSessionSavedItem}>
                        {name}
                      </Text>
                    ))}
                  </View>
                ) : null}
              </View>

              {!pendingPhotoLabel ? (
                <View style={styles.card}>
                  <Text style={styles.section}>Choose Photo Label</Text>
                  <Text style={styles.helper}>
                    Pick the label first, take the photo, and the app will bring you right back here for the next one.
                  </Text>
                  <View style={styles.photoLabelGrid}>
                    {PHOTO_LABEL_OPTIONS.map((label) => (
                      <Pressable
                        key={label}
                        onPress={() => startLabeledPhotoCapture(label)}
                        style={styles.photoLabelChip}
                      >
                        <Text style={styles.photoLabelChipText}>{label}</Text>
                      </Pressable>
                    ))}
                  </View>
                  <View style={styles.stack}>
                    <Action label="Finished" onPress={closePhotoSession} disabled={uploadingPhoto} />
                  </View>
                </View>
              ) : (
                <View style={styles.card}>
                  <Text style={styles.section}>Capture</Text>
                  <Text style={styles.helperStrong}>Current Label: {pendingPhotoLabel}</Text>
                  {cameraPermission?.granted ? (
                    <>
                      <Text style={styles.helper}>
                        Take the photo and it will save immediately under this label, then return to the label picker.
                      </Text>
                      <View style={styles.zoomRow}>
                        <Text style={styles.zoomLabel}>Zoom</Text>
                        <View style={styles.zoomChipRow}>
                          {PHOTO_ZOOM_OPTIONS.map((option) => (
                            <Pressable
                              key={option.label}
                              onPress={() => setPhotoZoom(option.value)}
                              style={[
                                styles.zoomChip,
                                photoZoom === option.value && styles.zoomChipActive,
                              ]}
                            >
                              <Text
                                style={[
                                  styles.zoomChipText,
                                  photoZoom === option.value && styles.zoomChipTextActive,
                                ]}
                              >
                                {option.label}
                              </Text>
                            </Pressable>
                          ))}
                        </View>
                      </View>
                      <View style={styles.cameraWrap}>
                        <CameraView
                          ref={cameraRef}
                          style={styles.cameraPreview}
                          facing="back"
                          mode="picture"
                          zoom={photoZoom}
                          autofocus={Platform.OS === "ios" ? "on" : undefined}
                          active={photoSessionVisible && !pendingPhotoAsset}
                          onCameraReady={() => setCameraReady(true)}
                          onMountError={(event) => handleError(new Error(event?.message || "Camera could not start."))}
                        />
                        {photoCaptureCueVisible ? (
                          <>
                            <View style={styles.cameraCaptureFlash} pointerEvents="none" />
                            <View style={styles.cameraCaptureBadgeWrap} pointerEvents="none">
                              <View style={styles.cameraCaptureBadge}>
                                <Text style={styles.cameraCaptureBadgeText}>
                                  {photoCaptureCueText || "Photo captured"}
                                </Text>
                              </View>
                            </View>
                          </>
                        ) : null}
                      </View>
                      <View style={styles.stack}>
                        <Action
                          label={
                            uploadingPhoto
                              ? "Saving..."
                              : !cameraReady
                                ? "Loading Camera..."
                                : "Take Photo"
                          }
                          primary
                          onPress={capturePhotoForSession}
                          disabled={uploadingPhoto || !cameraReady}
                        />
                        <Action label="Change Label" onPress={returnToPhotoLabelPicker} disabled={uploadingPhoto} />
                        <Action label="Finished" onPress={closePhotoSession} disabled={uploadingPhoto} />
                      </View>
                    </>
                  ) : (
                    <View style={styles.stack}>
                      <Text style={styles.helper}>
                        Camera permission is needed before you can take claim photos in the app.
                      </Text>
                      <Action label="Allow Camera" primary onPress={openPhotoSession} disabled={uploadingPhoto} />
                      <Action label="Finished" onPress={closePhotoSession} disabled={uploadingPhoto} />
                    </View>
                  )}
                </View>
              )}
            </ScrollView>
          </View>
        </SafeAreaView>
      </Modal>
    </SafeAreaView>
  );

  function renderContent() {
    if (homeTab === "claims") {
      if (claimTab === "new") {
        return (
          <ScrollView contentContainerStyle={styles.scrollPad}>
            <View style={styles.card}>
              <Text style={styles.section}>New Assignment</Text>
              <Text style={styles.helper}>
                Pick the assignment PDF from your phone. The Home PC will drop it into `PENDING CLAIMS`,
                run the same refresh-scan organizer, create the proper folder name, and populate the claim.
              </Text>
              <View style={styles.stack}>
                <Action
                  label={assignmentUploadInFlight ? "Importing..." : "Pick Assignment PDF"}
                  primary
                  onPress={uploadNewAssignmentPdf}
                  disabled={assignmentUploadInFlight}
                />
              </View>
            </View>
            {assignmentUploadResult ? (
              <View style={styles.card}>
                <Text style={styles.section}>Last Import</Text>
                <Text style={styles.helperStrong}>
                  {assignmentUploadResult.claim?.claim_id || "New claim imported"}
                </Text>
                <Text style={styles.helper}>
                  {(assignmentUploadResult.claim?.customer_name || assignmentUploadResult.claim?.title || "-") +
                    (assignmentUploadResult.folder_name
                      ? `\nFolder: ${assignmentUploadResult.folder_name}`
                      : "")}
                </Text>
                {assignmentUploadResult.claim?.key ? (
                  <Action
                    label="Open Imported Claim"
                    onPress={async () => {
                      setSelectedClaimKey(assignmentUploadResult.claim.key);
                      await loadClaimDetail(assignmentUploadResult.claim.key);
                      setDetailVisible(true);
                    }}
                  />
                ) : null}
              </View>
            ) : null}
          </ScrollView>
        );
      }
      return (
        <View style={styles.cardFill}>
          <FlatList
            data={claimRecords}
            keyExtractor={(item) => item.key}
            contentContainerStyle={styles.listContent}
            {...listPerfProps}
            renderItem={({ item }) => (
              <Pressable
                style={styles.listItem}
                onPress={() => {
                  setSelectedClaimKey(item.key);
                  setDetailVisible(true);
                }}
                onLongPress={() => promptPhoneActionsForRecord(item, item.customer_name || item.title)}
              >
                <Text style={styles.itemTitle}>
                  {[item.claim_id || "-", item.customer_name || item.title || "-"].join(BULLET)}
                </Text>
                {item.needs_measurement_photos ? (
                  <Text style={styles.measurementFlag}>MEASUREMENT PHOTOS REQUIRED</Text>
                ) : null}
                <Text style={styles.itemMeta}>{item.insurance_company || "-"}</Text>
                <Text style={styles.itemMeta}>
                  {[item.vehicle, item.shop_name, item.town].filter(Boolean).join(BULLET) || "-"}
                </Text>
                <Text style={[styles.itemMeta, styles.phoneMeta]}>
                  Phone: {formatPhone(resolvedPhoneForRecord(item)) || "No phone number listed"}
                </Text>
              </Pressable>
            )}
            ItemSeparatorComponent={() => <View style={styles.sep} />}
          />
        </View>
      );
    }

    if (homeTab === "route") {
      return (
        <ScrollView contentContainerStyle={styles.scrollPad}>
          <View style={styles.card}>
            <Text style={styles.section}>Route Planner</Text>
            <Text style={styles.helper}>
              Choose any claims, move them into your route, fix any address that looks off, reorder the stops, then open the trip in your navigation app.
            </Text>
            <Text style={[styles.section, styles.spaced]}>Start From</Text>
            <Text style={styles.itemMeta}>
              {routeData.start_from || dashboard.settings?.route_home_address || "-"}
            </Text>
            <View style={[styles.tabRow, styles.spaced]}>
              {ROUTE_STATUS_TABS.map(([key, label]) => (
                <Chip key={key} label={label} active={routeStatus === key} onPress={() => setRouteStatus(key)} />
              ))}
            </View>
            <View style={styles.routeSummaryRow}>
              <View style={styles.routeSummaryCard}>
                <Text style={styles.summaryLabel}>Selected Stops</Text>
                <Text style={styles.summaryValue}>{plannedRouteRecords.length}</Text>
              </View>
              <Action
                label={routeSaving ? "..." : "Clear"}
                compact
                disabled={routeSaving || !plannedRouteRecords.length}
                onPress={clearRoutePlan}
              />
            </View>
            <View style={styles.routeLaunchRow}>
              <Action
                label={routeSaving ? "Saving..." : "Google Maps"}
                primary
                compact
                disabled={routeSaving}
                onPress={openGoogleMapsRoute}
              />
              <Action
                label="Waze"
                compact
                disabled={routeSaving || !plannedRouteRecords.length}
                onPress={openWazeRoute}
              />
              <Action
                label="Apple Maps"
                compact
                disabled={routeSaving || !plannedRouteRecords.length}
                onPress={openAppleMapsRoute}
              />
            </View>
            <Text style={[styles.section, styles.spaced]}>Selected Route</Text>
            {plannedRouteRecords.length ? (
              plannedRouteRecords.map((item, index) => (
                <RouteRow
                  key={`planned-${item.key}`}
                  item={item}
                  badge={`Stop ${index + 1}`}
                  onPress={() => openRouteClaim(item.key)}
                  onPhonePress={() => promptPhoneActionsForRecord(item, item.customer_name)}
                  actions={[
                    {
                      label: "Up",
                      onPress: () => moveRouteStop(item.key, -1),
                      disabled: routeSaving || index === 0,
                    },
                    {
                      label: "Down",
                      onPress: () => moveRouteStop(item.key, 1),
                      disabled: routeSaving || index === plannedRouteRecords.length - 1,
                    },
                    {
                      label: "Edit Address",
                      onPress: () => startEditingRouteAddress(item),
                      disabled: routeSaving || routeAddressSaving,
                    },
                    {
                      label: "Remove",
                      onPress: () => removeRouteStop(item.key),
                      disabled: routeSaving,
                      tone: "danger",
                    },
                  ]}
                  editor={
                    editingRouteKey === item.key ? (
                      <View style={styles.routeEditor}>
                        <TextInput
                          value={editingRouteAddress}
                          onChangeText={setEditingRouteAddress}
                          autoCapitalize="words"
                          autoCorrect={false}
                          style={styles.routeEditorInput}
                          placeholder="Edit route address..."
                        />
                        <View style={styles.routeEditorActions}>
                          <Action
                            label={routeAddressSaving ? "Saving..." : "Save Address"}
                            primary
                            compact
                            disabled={routeAddressSaving}
                            onPress={saveEditedRouteAddress}
                          />
                          <Action
                            label="Cancel"
                            compact
                            disabled={routeAddressSaving}
                            onPress={cancelEditingRouteAddress}
                          />
                        </View>
                      </View>
                    ) : null
                  }
                />
              ))
            ) : (
              <Text style={styles.helper}>No claims selected yet. Add claims from the list below.</Text>
            )}
            <Text style={[styles.section, styles.spaced]}>Navigation</Text>
            <Text style={styles.helper}>
              Google Maps opens the full multi-stop route. Waze and Apple Maps open the last selected stop as the destination.
            </Text>
            <Text style={[styles.section, styles.spaced]}>Available Claims</Text>
            {availableRouteRecords.length ? (
              availableRouteRecords.map((item) => (
                <RouteRow
                  key={item.key}
                  item={item}
                  onPress={() => openRouteClaim(item.key)}
                  onPhonePress={() => promptPhoneActionsForRecord(item, item.customer_name)}
                  actions={[
                    {
                      label: "Add",
                      onPress: () => addRouteStop(item),
                      disabled: routeSaving,
                      tone: "primary",
                    },
                  ]}
                />
              ))
            ) : (
              <Text style={styles.helper}>No claims match this filter right now.</Text>
            )}
          </View>
        </ScrollView>
      );
    }

    if (libraryTab === "tools") {
      return (
        <ScrollView contentContainerStyle={styles.scrollPad}>
          <View style={styles.card}>
            <Text style={styles.section}>Claim Tools</Text>
            <Text style={styles.itemMeta}>{toolsData.claim_tools_folder || "Claim tools folder not set."}</Text>
            <Text style={[styles.section, styles.spaced]}>Contacts</Text>
            {(toolsData.contacts || []).map((entry, index) => (
              <LibraryRow
                key={`contact-${index}`}
                title={entry.name}
                lines={[
                  `Number: ${entry.number || "-"}`,
                  `Prompt Guide: ${entry.prompt_guide || "-"}`,
                  entry.notes || "No notes.",
                ]}
              />
            ))}
            <Text style={[styles.section, styles.spaced]}>Files</Text>
            {(toolsData.files || []).map((entry, index) => (
              <LibraryRow
                key={`file-${index}`}
                title={entry.label || entry.file_name}
                lines={[`Section: ${entry.section || "-"}`, entry.notes || "No notes."]}
              />
            ))}
          </View>
        </ScrollView>
      );
    }

    if (libraryTab === "shops") {
      return (
      <FlatList
        data={bodyShops}
        keyExtractor={(item, index) => `${item.shop_name}-${index}`}
        contentContainerStyle={styles.listContent}
        {...listPerfProps}
        renderItem={({ item }) => (
            <LibraryRow
              title={item.shop_name}
              lines={[
                `Contact: ${item.contact_name || "-"}`,
                `Phone: ${item.phone || "-"}`,
                `Email: ${item.email || "-"}`,
                `Tax ID: ${item.tax_id || "-"}`,
                item.notes || item.negotiation_notes || "No notes.",
              ]}
            />
          )}
          ItemSeparatorComponent={() => <View style={styles.sep} />}
        />
      );
    }

    return (
      <FlatList
        data={insuranceCompanies}
        keyExtractor={(item, index) => `${item.company_name}-${index}`}
        contentContainerStyle={styles.listContent}
        {...listPerfProps}
        renderItem={({ item }) => (
          <LibraryRow
            title={item.company_name}
            lines={[
              `Quick Summary: ${item.quick_summary || "-"}`,
              `Labor Rates: ${item.labor_rates || "-"}`,
              `Total Loss Threshold: ${item.total_loss_threshold || "-"}`,
              item.photo_rules || item.documentation_requirements || item.notes || "No notes.",
            ]}
          />
        )}
        ItemSeparatorComponent={() => <View style={styles.sep} />}
      />
    );
  }
}

function Summary({ label, value }) {
  return (
    <View style={styles.summary}>
      <Text style={styles.summaryLabel}>{label}</Text>
      <Text style={styles.summaryValue}>{value ?? 0}</Text>
    </View>
  );
}

function Chip({ label, active, onPress }) {
  return (
    <Pressable onPress={onPress} style={[styles.chip, active && styles.chipOn]}>
      <Text style={[styles.chipText, active && styles.chipTextOn]}>{label}</Text>
    </Pressable>
  );
}

function FieldRow({ label, value, onPress, onLongPress }) {
  const content = (
    <>
      <Text style={styles.fieldLabel}>{label}</Text>
      <Text style={[styles.fieldValue, (onPress || onLongPress) && styles.fieldValueAction]}>{value || "-"}</Text>
    </>
  );
  if (onPress || onLongPress) {
    return (
      <Pressable style={styles.fieldRow} onPress={onPress} onLongPress={onLongPress}>
        {content}
      </Pressable>
    );
  }
  return <View style={styles.fieldRow}>{content}</View>;
}

function Action({ label, primary, onPress, compact, disabled }) {
  return (
    <Pressable
      onPress={onPress}
      disabled={disabled}
      style={[
        styles.action,
        compact && styles.actionCompact,
        primary && styles.actionPrimary,
        disabled && styles.actionDisabled,
      ]}
    >
      <Text style={[styles.actionText, primary && styles.actionTextPrimary, disabled && styles.actionTextDisabled]}>
        {label}
      </Text>
    </Pressable>
  );
}

function RouteRow({ item, onPress, onPhonePress, badge, actions = [], editor = null }) {
  return (
    <View style={styles.listItem}>
      {badge ? <Text style={styles.routeBadge}>{badge}</Text> : null}
      <Pressable onPress={onPress} onLongPress={onPhonePress}>
        <Text style={styles.itemTitle}>{[item.claim_id || "-", item.customer_name || "-"].join(BULLET)}</Text>
        {item.needs_measurement_photos ? (
          <Text style={styles.measurementFlag}>MEASUREMENT PHOTOS REQUIRED</Text>
        ) : null}
        <Text style={styles.itemMeta}>{item.address || "-"}</Text>
        <Text style={styles.itemMeta}>
          {[item.status_label || item.status, item.shop_name, item.insurance_company, item.vehicle]
            .filter(Boolean)
            .join(BULLET) || "-"}
        </Text>
        <Text style={[styles.itemMeta, styles.phoneMeta]}>
          Phone: {formatPhone(item.contact_phone) || "No phone number listed"}
        </Text>
      </Pressable>
      {actions.length ? (
        <View style={styles.routeActions}>
          {actions.map((action) => (
            <Pressable
              key={action.label}
              onPress={action.onPress}
              disabled={action.disabled}
              style={[
                styles.routeAction,
                action.tone === "primary" && styles.routeActionPrimary,
                action.tone === "danger" && styles.routeActionDanger,
                action.disabled && styles.routeActionDisabled,
              ]}
            >
              <Text
                style={[
                  styles.routeActionText,
                  action.tone === "primary" && styles.routeActionTextPrimary,
                  action.tone === "danger" && styles.routeActionTextDanger,
                  action.disabled && styles.routeActionTextDisabled,
                ]}
              >
                {action.label}
              </Text>
            </Pressable>
          ))}
        </View>
      ) : null}
      {editor}
    </View>
  );
}

function LibraryRow({ title, lines }) {
  return (
    <View style={styles.listItem}>
      {title ? <Text style={styles.itemTitle}>{title}</Text> : null}
      {lines.filter(Boolean).map((line, idx) => (
        <Text key={idx} style={styles.itemMeta}>
          {line}
        </Text>
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  safe: {
    flex: 1,
    backgroundColor: "#eef4fb",
  },
  shell: {
    flex: 1,
    paddingHorizontal: scaleUi(12),
    paddingTop: Platform.OS === "android" ? (NativeStatusBar.currentHeight || 0) + scaleUi(16) : scaleUi(10),
    paddingBottom: 0,
    gap: scaleUi(6),
  },
  headerTop: {
    flexDirection: "row",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: scaleUi(8),
  },
  header: {
    flex: 1,
    gap: 0,
  },
  eyebrow: {
    fontSize: scaleUi(9),
    color: "#6a7f9d",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: 1.2,
  },
  title: {
    fontSize: scaleUi(15),
    fontWeight: "800",
    color: "#0f2742",
  },
  subtitle: {
    fontSize: scaleUi(10),
    color: "#5d718d",
    lineHeight: scaleUi(14),
  },
  versionText: {
    marginTop: scaleUi(4),
    fontSize: scaleUi(11),
    fontWeight: "800",
    color: "#b42318",
  },
  versionSubtext: {
    fontSize: scaleUi(9),
    color: "#6a7f9d",
    fontWeight: "700",
  },
  card: {
    backgroundColor: "#ffffff",
    borderRadius: scaleUi(18),
    padding: scaleUi(14),
    borderWidth: 1,
    borderColor: "#d7e1ee",
    gap: scaleUi(8),
  },
  utilityCard: {
    backgroundColor: "#ffffff",
    borderRadius: scaleUi(16),
    padding: scaleUi(10),
    borderWidth: 1,
    borderColor: "#d7e1ee",
    gap: scaleUi(6),
  },
  statusRow: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    minHeight: 24,
  },
  statusChip: {
    flexDirection: "row",
    alignItems: "center",
    gap: scaleUi(6),
    backgroundColor: "#ffffff",
    borderRadius: 999,
    paddingHorizontal: scaleUi(9),
    paddingVertical: scaleUi(4),
    borderWidth: 1,
    borderColor: "#d7e1ee",
    maxWidth: scaleUi(190),
  },
  pendingStatusChip: {
    maxWidth: scaleUi(168),
  },
  contentWrap: {
    flex: 1,
  },
  statusChipLabel: {
    fontSize: scaleUi(9),
    fontWeight: "800",
    color: "#6b809d",
    textTransform: "uppercase",
    letterSpacing: 0.8,
  },
  statusChipValue: {
    flex: 1,
    fontSize: scaleUi(10),
    color: "#183659",
    fontWeight: "600",
  },
  statusMessage: {
    flex: 1,
    fontSize: scaleUi(10),
    color: "#6b809d",
  },
  utilityHeader: {
    flexDirection: "row",
    alignItems: "center",
    gap: 10,
  },
  utilityHeaderText: {
    flex: 1,
    gap: 2,
  },
  utilityTitle: {
    fontSize: 13,
    fontWeight: "800",
    color: "#183659",
    textTransform: "uppercase",
    letterSpacing: 0.8,
  },
  utilityHint: {
    fontSize: 12,
    color: "#6d829f",
    lineHeight: 16,
  },
  sheetBackdrop: {
    flex: 1,
    backgroundColor: "rgba(11, 23, 40, 0.24)",
    justifyContent: "flex-end",
  },
  sheetDismiss: {
    flex: 1,
  },
  sheetCard: {
    backgroundColor: "#eef4fb",
    borderTopLeftRadius: 24,
    borderTopRightRadius: 24,
    paddingHorizontal: 16,
    paddingTop: 12,
    paddingBottom: 18,
    gap: 10,
    borderTopWidth: 1,
    borderColor: "#d7e1ee",
  },
  cardFill: {
    flex: 1,
    backgroundColor: "#ffffff",
    borderRadius: 20,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    overflow: "hidden",
  },
  summaryRow: {
    flexDirection: "row",
    gap: 6,
  },
  summary: {
    flex: 1,
    backgroundColor: "#ffffff",
    borderRadius: 14,
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderWidth: 1,
    borderColor: "#d7e1ee",
  },
  summaryLabel: {
    fontSize: 10,
    color: "#6e84a3",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: 0.7,
  },
  summaryValue: {
    marginTop: 1,
    fontSize: 15,
    fontWeight: "800",
    color: "#0f2742",
  },
  label: {
    fontSize: 13,
    color: "#6e84a3",
    fontWeight: "700",
  },
  helper: {
    fontSize: scaleUi(12),
    color: "#6e84a3",
    lineHeight: scaleUi(16),
  },
  helperStrong: {
    fontSize: scaleUi(11),
    fontWeight: "800",
    color: "#183659",
  },
  input: {
    borderWidth: 1,
    borderColor: "#cad7e8",
    backgroundColor: "#fbfdff",
    borderRadius: scaleUi(12),
    paddingHorizontal: scaleUi(12),
    paddingVertical: scaleUi(9),
    fontSize: scaleUi(13),
    color: "#10253f",
  },
  search: {
    borderWidth: 1,
    borderColor: "#cad7e8",
    backgroundColor: "#ffffff",
    borderRadius: scaleUi(13),
    paddingHorizontal: scaleUi(13),
    paddingVertical: scaleUi(8),
    fontSize: scaleUi(12),
    color: "#10253f",
  },
  tabRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 5,
  },
  chip: {
    minWidth: scaleUi(62),
    paddingHorizontal: scaleUi(10),
    paddingVertical: scaleUi(5),
    borderRadius: 999,
    backgroundColor: "#e7eef7",
    alignItems: "center",
  },
  chipOn: {
    backgroundColor: "#103c6d",
  },
  chipText: {
    color: "#35506f",
    fontSize: scaleUi(10),
    fontWeight: "700",
  },
  chipTextOn: {
    color: "#ffffff",
  },
  loading: {
    flex: 1,
    alignItems: "center",
    justifyContent: "center",
    gap: 12,
  },
  loadingOverlay: {
    position: "absolute",
    top: 6,
    left: 0,
    right: 0,
    alignItems: "center",
  },
  loadingPill: {
    flexDirection: "row",
    alignItems: "center",
    gap: 8,
    backgroundColor: "rgba(255,255,255,0.98)",
    borderRadius: 999,
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderWidth: 1,
    borderColor: "#d7e1ee",
  },
  loadingPillText: {
    fontSize: scaleUi(11),
    fontWeight: "700",
    color: "#35506f",
  },
  error: {
    backgroundColor: "#feecec",
    borderRadius: 14,
    padding: 12,
    borderWidth: 1,
    borderColor: "#f4b8b8",
  },
  errorText: {
    color: "#8f2c2c",
    fontSize: scaleUi(12),
    lineHeight: scaleUi(18),
  },
  listItem: {
    marginHorizontal: scaleUi(8),
    marginVertical: scaleUi(3),
    paddingHorizontal: scaleUi(11),
    paddingVertical: scaleUi(8),
    backgroundColor: "#ffffff",
    borderRadius: scaleUi(14),
    borderWidth: 1,
    borderColor: "#e1e9f3",
  },
  fileListItem: {
    flexDirection: "row",
    alignItems: "center",
    gap: 10,
  },
  fileMainArea: {
    flex: 1,
  },
  fileActionColumn: {
    alignItems: "flex-end",
    gap: 8,
  },
  fileSelectChip: {
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 8,
    backgroundColor: "#edf3fa",
    borderWidth: 1,
    borderColor: "#d5e0ee",
  },
  fileSelectChipActive: {
    backgroundColor: "#103c6d",
    borderColor: "#103c6d",
  },
  fileSelectChipText: {
    fontSize: 11,
    fontWeight: "800",
    color: "#17395d",
  },
  fileSelectChipTextActive: {
    color: "#ffffff",
  },
  fileNonPdfHint: {
    fontSize: 11,
    fontWeight: "700",
    color: "#6e84a3",
  },
  fileDeleteChip: {
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 8,
    backgroundColor: "#fdecec",
    borderWidth: 1,
    borderColor: "#f2bbbb",
  },
  fileDeleteChipText: {
    fontSize: 11,
    fontWeight: "800",
    color: "#a03333",
  },
  itemTitle: {
    fontSize: scaleUi(12.5),
    fontWeight: "700",
    color: "#10253f",
  },
  measurementFlag: {
    alignSelf: "flex-start",
    marginTop: 6,
    backgroundColor: "#fff0ef",
    color: "#bf1d1d",
    fontSize: 11,
    fontWeight: "900",
    letterSpacing: 0.4,
    paddingHorizontal: 8,
    paddingVertical: 4,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: "#f4b8b8",
  },
  itemMeta: {
    marginTop: scaleUi(3),
    fontSize: scaleUi(10),
    lineHeight: scaleUi(14),
    color: "#5d718d",
  },
  phoneMeta: {
    color: "#0e3a66",
    fontWeight: "800",
  },
  sep: {
    height: 0,
  },
  modalHead: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
  },
  modalTitle: {
    fontSize: scaleUi(17),
    fontWeight: "800",
    color: "#10253f",
  },
  smallBtn: {
    backgroundColor: "#e8eef7",
    borderRadius: scaleUi(12),
    paddingHorizontal: scaleUi(12),
    paddingVertical: scaleUi(8),
  },
  smallBtnText: {
    color: "#14385f",
    fontWeight: "700",
    fontSize: scaleUi(12),
  },
  scrollPad: {
    paddingBottom: 24,
    gap: 12,
  },
  section: {
    fontSize: scaleUi(17),
    fontWeight: "800",
    color: "#10253f",
  },
  routeSummaryRow: {
    flexDirection: "row",
    alignItems: "stretch",
    gap: 8,
    marginTop: 8,
  },
  routeLaunchRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
    marginTop: 8,
  },
  routeSummaryCard: {
    minWidth: 92,
    backgroundColor: "#f4f8fd",
    borderRadius: 16,
    paddingHorizontal: 12,
    paddingVertical: 10,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    justifyContent: "center",
  },
  spaced: {
    marginTop: 8,
  },
  fieldRow: {
    borderBottomWidth: 1,
    borderBottomColor: "#dce5f0",
    paddingBottom: 8,
    marginBottom: 2,
  },
  fieldLabel: {
    fontSize: scaleUi(11),
    color: "#6e84a3",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: 0.8,
  },
  fieldValue: {
    marginTop: scaleUi(3),
    fontSize: scaleUi(15),
    color: "#10253f",
    lineHeight: scaleUi(20),
  },
  fieldValueAction: {
    color: "#0e3a66",
    fontWeight: "800",
  },
  stack: {
    gap: 10,
  },
  action: {
    backgroundColor: "#e8eef7",
    borderRadius: scaleUi(14),
    paddingVertical: scaleUi(11),
    paddingHorizontal: scaleUi(14),
    alignItems: "center",
  },
  actionCompact: {
    flex: 1,
    paddingVertical: scaleUi(9),
    paddingHorizontal: scaleUi(10),
  },
  actionPrimary: {
    backgroundColor: "#103c6d",
  },
  actionDisabled: {
    opacity: 0.55,
  },
  actionText: {
    fontSize: scaleUi(14),
    fontWeight: "700",
    color: "#17395d",
  },
  actionTextPrimary: {
    color: "#ffffff",
  },
  actionTextDisabled: {
    color: "#5d718d",
  },
  noteBlock: {
    fontSize: 14,
    lineHeight: 21,
    color: "#445a77",
    backgroundColor: "#f6f9fc",
    borderRadius: 14,
    padding: 14,
  },
  notesInput: {
    minHeight: 100,
    borderWidth: 1,
    borderColor: "#cad7e8",
    backgroundColor: "#fbfdff",
    borderRadius: 14,
    paddingHorizontal: 14,
    paddingVertical: 12,
    textAlignVertical: "top",
    fontSize: 16,
    color: "#10253f",
  },
  emailSelectionList: {
    gap: 6,
    maxHeight: 180,
  },
  recipientPickerWrap: {
    gap: 8,
  },
  recipientPickerToggle: {
    alignSelf: "flex-start",
    backgroundColor: "#edf3fa",
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 8,
    borderWidth: 1,
    borderColor: "#d5e0ee",
  },
  recipientPickerToggleText: {
    fontSize: 11,
    fontWeight: "800",
    color: "#17395d",
  },
  recipientPickerMenu: {
    gap: 8,
    backgroundColor: "#f6f9fc",
    borderRadius: 14,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    paddingHorizontal: 10,
    paddingVertical: 10,
  },
  recipientPickerRow: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
  },
  recipientPickerInfo: {
    flex: 1,
    gap: 2,
  },
  recipientPickerName: {
    fontSize: 12,
    fontWeight: "800",
    color: "#10253f",
  },
  recipientPickerEmail: {
    fontSize: 11,
    color: "#5d718d",
  },
  recipientPickerActions: {
    flexDirection: "row",
    gap: 8,
  },
  recipientPickerChip: {
    minWidth: 40,
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: "#ffffff",
    borderWidth: 1,
    borderColor: "#d5e0ee",
    borderRadius: 999,
    paddingHorizontal: 10,
    paddingVertical: 7,
  },
  recipientPickerChipText: {
    fontSize: 11,
    fontWeight: "800",
    color: "#17395d",
  },
  emailSelectionItem: {
    fontSize: 12,
    lineHeight: 17,
    color: "#445a77",
    backgroundColor: "#f6f9fc",
    borderRadius: 10,
    paddingHorizontal: 10,
    paddingVertical: 8,
  },
  updateBtn: {
    backgroundColor: "#103c6d",
    borderRadius: 12,
    paddingVertical: 7,
    paddingHorizontal: 10,
    alignItems: "center",
    justifyContent: "center",
    minWidth: 58,
  },
  updateBtnText: {
    color: "#ffffff",
    fontSize: 11,
    fontWeight: "700",
    textAlign: "center",
  },
  ghostBtn: {
    backgroundColor: "#edf3fa",
    borderRadius: 12,
    paddingVertical: 7,
    paddingHorizontal: 10,
    alignItems: "center",
    justifyContent: "center",
  },
  ghostBtnText: {
    color: "#183659",
    fontSize: 11,
    fontWeight: "700",
  },
  listContent: {
    paddingVertical: 4,
    paddingBottom: 18,
  },
  routeBadge: {
    alignSelf: "flex-start",
    backgroundColor: "#edf3fa",
    color: "#17395d",
    fontSize: 10,
    fontWeight: "800",
    textTransform: "uppercase",
    letterSpacing: 0.7,
    borderRadius: 999,
    paddingHorizontal: 8,
    paddingVertical: 4,
    marginBottom: 8,
  },
  routeActions: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
    marginTop: 10,
  },
  routeAction: {
    borderRadius: 999,
    paddingHorizontal: 10,
    paddingVertical: 6,
    backgroundColor: "#edf3fa",
  },
  routeActionPrimary: {
    backgroundColor: "#103c6d",
  },
  routeActionDanger: {
    backgroundColor: "#fce8e8",
  },
  routeActionDisabled: {
    opacity: 0.5,
  },
  routeActionText: {
    fontSize: 11,
    fontWeight: "800",
    color: "#17395d",
  },
  routeActionTextPrimary: {
    color: "#ffffff",
  },
  routeActionTextDanger: {
    color: "#8f2c2c",
  },
  routeActionTextDisabled: {
    color: "#5d718d",
  },
  inlineHint: {
    fontSize: 11,
    color: "#6b809d",
    marginTop: -2,
    marginBottom: 2,
  },
  photoRow: {
    flexDirection: "row",
    gap: 10,
  },
  queueNotice: {
    gap: 8,
    backgroundColor: "#f6f9fc",
    borderRadius: 14,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    paddingHorizontal: 12,
    paddingVertical: 10,
  },
  queueNoticeTitle: {
    fontSize: 12,
    fontWeight: "800",
    color: "#183659",
  },
  photoPreview: {
    width: "100%",
    aspectRatio: 3 / 4,
    borderRadius: 18,
    backgroundColor: "#dce5f0",
  },
  cameraWrap: {
    overflow: "hidden",
    borderRadius: 18,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    backgroundColor: "#0c1623",
    position: "relative",
  },
  cameraPreview: {
    width: "100%",
    aspectRatio: 3 / 4,
  },
  captureCueBanner: {
    borderRadius: 14,
    paddingHorizontal: 12,
    paddingVertical: 10,
    backgroundColor: "#e7f0fb",
    borderWidth: 1,
    borderColor: "#b9d0ea",
  },
  captureCueText: {
    fontSize: 13,
    fontWeight: "900",
    color: "#103c6d",
    textAlign: "center",
    textTransform: "uppercase",
    letterSpacing: 0.7,
  },
  cameraCaptureFlash: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: "rgba(255,255,255,0.7)",
  },
  cameraCaptureBadgeWrap: {
    ...StyleSheet.absoluteFillObject,
    alignItems: "center",
    justifyContent: "center",
    paddingHorizontal: 18,
  },
  cameraCaptureBadge: {
    borderRadius: 18,
    paddingHorizontal: 18,
    paddingVertical: 14,
    backgroundColor: "rgba(11, 23, 40, 0.86)",
    borderWidth: 1,
    borderColor: "rgba(255,255,255,0.32)",
  },
  cameraCaptureBadgeText: {
    fontSize: 16,
    fontWeight: "900",
    color: "#ffffff",
    textAlign: "center",
    textTransform: "uppercase",
    letterSpacing: 1,
  },
  photoLabelGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  photoLabelChip: {
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 9,
    backgroundColor: "#edf3fa",
    borderWidth: 1,
    borderColor: "#d5e0ee",
  },
  photoLabelChipActive: {
    backgroundColor: "#103c6d",
    borderColor: "#103c6d",
  },
  photoLabelChipText: {
    fontSize: 12,
    fontWeight: "800",
    color: "#17395d",
  },
  photoLabelChipTextActive: {
    color: "#ffffff",
  },
  zoomRow: {
    gap: 8,
  },
  zoomLabel: {
    fontSize: 12,
    fontWeight: "800",
    letterSpacing: 0.8,
    color: "#577291",
    textTransform: "uppercase",
  },
  zoomChipRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  zoomChip: {
    borderRadius: 999,
    paddingHorizontal: 12,
    paddingVertical: 9,
    backgroundColor: "#edf3fa",
    borderWidth: 1,
    borderColor: "#d5e0ee",
  },
  zoomChipActive: {
    backgroundColor: "#103c6d",
    borderColor: "#103c6d",
  },
  zoomChipText: {
    fontSize: 12,
    fontWeight: "800",
    color: "#17395d",
  },
  zoomChipTextActive: {
    color: "#ffffff",
  },
  photoSessionSavedList: {
    gap: 6,
  },
  photoSessionSavedItem: {
    fontSize: 12,
    lineHeight: 17,
    color: "#445a77",
    backgroundColor: "#f6f9fc",
    borderRadius: 10,
    paddingHorizontal: 10,
    paddingVertical: 8,
  },
  routeEditor: {
    marginTop: 10,
    gap: 8,
  },
  routeEditorInput: {
    borderWidth: 1,
    borderColor: "#cad7e8",
    backgroundColor: "#fbfdff",
    borderRadius: 12,
    paddingHorizontal: 12,
    paddingVertical: 10,
    fontSize: 14,
    color: "#10253f",
  },
  routeEditorActions: {
    flexDirection: "row",
    gap: 10,
  },
});
