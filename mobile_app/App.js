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
  FlatList,
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
  View,
} from "react-native";
import { StatusBar as ExpoStatusBar } from "expo-status-bar";
import * as ImagePicker from "expo-image-picker";
import * as ImageManipulator from "expo-image-manipulator";
import * as Updates from "expo-updates";
import appConfig from "./app.json";

const API_BASE_DEFAULT = "https://api.luxuryimportsusa.shop";
const BULLET = " | ";
const API_BASE_STORAGE_KEY = "claim_manager_mobile_api_base";
const ROUTE_PLAN_STORAGE_KEY = "claim_manager_mobile_route_plan_keys";
const ROUTE_ADDRESS_OVERRIDES_STORAGE_KEY = "claim_manager_mobile_route_address_overrides";
const APP_VERSION = appConfig?.expo?.version || "unknown";

const HOME_TABS = [
  ["claims", "Claims"],
  ["route", "Route Planner"],
  ["library", "Library"],
];

const CLAIM_TABS = [
  ["open", "Open"],
  ["closed", "Closed"],
  ["all", "All"],
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

async function prepareUploadPhoto(asset) {
  const manipulations = [];
  const width = asset.width || 0;
  const height = asset.height || 0;
  const longestSide = Math.max(width, height);

  if (longestSide > 1280) {
    if (width >= height) {
      manipulations.push({ resize: { width: 1280 } });
    } else {
      manipulations.push({ resize: { height: 1280 } });
    }
  }

  let result = await ImageManipulator.manipulateAsync(
    asset.uri,
    manipulations,
    {
      compress: 0.28,
      format: ImageManipulator.SaveFormat.JPEG,
      base64: false,
    }
  );

  if ((asset.fileSize || 0) > 500000) {
    result = await ImageManipulator.manipulateAsync(
      result.uri,
      [],
      {
        compress: 0.18,
        format: ImageManipulator.SaveFormat.JPEG,
        base64: false,
      }
    );
  }

  return {
    uri: result.uri,
    name: (asset.fileName || "claim-photo").replace(/\.[^.]+$/, "") + ".jpg",
    type: "image/jpeg",
  };
}

export default function App() {
  const [apiBase, setApiBase] = useState(API_BASE_DEFAULT);
  const [utilityOpen, setUtilityOpen] = useState(false);
  const [homeTab, setHomeTab] = useState("claims");
  const [claimTab, setClaimTab] = useState("open");
  const [routeStatus, setRouteStatus] = useState("all");
  const [libraryTab, setLibraryTab] = useState("tools");
  const [search, setSearch] = useState("");
  const [dashboard, setDashboard] = useState({
    summary: { all: 0, open: 0, closed: 0 },
    settings: {},
  });
  const [claimRecords, setClaimRecords] = useState([]);
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
  const deferredSearch = useDeferredValue(search);
  const currentChannel = Updates.channel || "preview";
  const currentRuntimeVersion = Updates.runtimeVersion || APP_VERSION;
  const currentUpdateId = Updates.updateId ? Updates.updateId.slice(0, 8) : "embedded";
  const currentVersionLabel = `v${APP_VERSION} / ${currentChannel} / ${currentUpdateId}`;
  const currentVersionDetail = `Runtime ${currentRuntimeVersion}`;
  const listPerfProps = {
    initialNumToRender: 10,
    maxToRenderPerBatch: 10,
    windowSize: 7,
    removeClippedSubviews: true,
  };

  useEffect(() => {
    let cancelled = false;

    async function loadStoredApiBase() {
      try {
        const saved = (await AsyncStorage.getItem(API_BASE_STORAGE_KEY))?.trim();
        if (!cancelled && saved) {
          setApiBase(saved);
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
    AsyncStorage.setItem(API_BASE_STORAGE_KEY, apiBase.trim()).catch(() => {});
  }, [apiBase, apiReady]);

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

  async function fetchJson(path, options) {
    const response = await fetch(`${apiBase}${path}`, options);
    if (!response.ok) {
      let message = `Request failed (${response.status}).`;
      try {
        const payload = await response.json();
        if (payload?.error) {
          message = payload.error;
        }
      } catch {}
      throw new Error(message);
    }
    return response.json();
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
        const [routePayload, claimsPayload] = await Promise.all([
          fetchJson("/api/mobile/route-planner?status=all"),
          fetchJson("/api/mobile/claims?status=all"),
        ]);
        if (requestId !== loadSequence.current) {
          return;
        }
        const claimStatusByKey = new Map(
          (claimsPayload.records || []).map((record) => [record.key, record.status])
        );
        const mergeRouteStatus = (record) => ({
          ...record,
          status: claimStatusByKey.get(record.key) || record.status,
        });
        startTransition(() => {
          setRouteData({
            ...routePayload,
            records: (routePayload.records || []).map(mergeRouteStatus),
            planned_records: (routePayload.planned_records || []).map(mergeRouteStatus),
          });
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
      setNoteText("");
    } catch (err) {
      handleError(err);
    } finally {
      setDetailLoading(false);
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

  async function uploadPhoto(useCamera) {
    if (!selectedClaimKey) {
      return;
    }
    setUploadingPhoto(true);
    try {
      const permission = useCamera
        ? await ImagePicker.requestCameraPermissionsAsync()
        : await ImagePicker.requestMediaLibraryPermissionsAsync();
      if (!permission.granted) {
        throw new Error(
          useCamera ? "Camera permission is required." : "Photo library permission is required."
        );
      }

      const result = useCamera
        ? await ImagePicker.launchCameraAsync({ quality: 0.9 })
        : await ImagePicker.launchImageLibraryAsync({
            quality: 0.9,
            allowsMultipleSelection: true,
            selectionLimit: 0,
          });

      if (result.canceled || !result.assets?.length) {
        return;
      }

      for (const asset of result.assets) {
        const uploadAsset = await prepareUploadPhoto(asset);
        const formData = new FormData();
        formData.append("photo", {
          uri: uploadAsset.uri,
          name: uploadAsset.name,
          type: uploadAsset.type,
        });

        await fetchJson(`/api/mobile/upload-photo?key=${encodeURIComponent(selectedClaimKey)}`, {
          method: "POST",
          body: formData,
        });
      }

      setUpdateMessage(
        result.assets.length > 1 ? `${result.assets.length} photos uploaded.` : "Photo uploaded."
      );
      await loadClaimDetail(selectedClaimKey);
    } catch (err) {
      handleError(err);
    } finally {
      setUploadingPhoto(false);
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
      ["Claim ID", claimDetail.claim_id],
      ["Customer", claimDetail.customer_name || claimDetail.title],
      ["Insurance", claimDetail.insurance_company],
      ["Claim #", claimDetail.claim_number],
      ["Status", claimDetail.status],
      ["Type", claimDetail.claim_type],
      ["Date Of Loss", claimDetail.date_of_loss],
      ["Town", claimDetail.town],
      ["Shop", claimDetail.shop_name],
      ["Vehicle", claimDetail.vehicle],
      ["VIN", claimDetail.vin],
      ["Inspection Location", claimDetail.route_display_address],
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
      <ExpoStatusBar style="dark" />
      <View style={styles.shell}>
        <View style={styles.headerTop}>
          <View style={styles.header}>
            <Text style={styles.eyebrow}>Fernando Marin</Text>
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

        {homeTab === "claims" ? (
          <>
            <View style={styles.tabRow}>
              {CLAIM_TABS.map(([key, label]) => (
                <Chip key={key} label={label} active={claimTab === key} onPress={() => setClaimTab(key)} />
              ))}
            </View>
            <Text style={styles.inlineHint}>
              Tap any claim to open files, photos, notes, and quick actions.
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
              <Text style={styles.utilityHint}>Use the secure Cloudflare address by default, or switch back to local if you are on the same network.</Text>
              <Text style={styles.helperStrong}>{currentVersionLabel}</Text>
              <Text style={styles.helper}>{currentVersionDetail}</Text>
              <TextInput
                value={apiBase}
                onChangeText={setApiBase}
                autoCapitalize="none"
                autoCorrect={false}
                style={styles.input}
                placeholder="https://api.luxuryimportsusa.shop"
              />
              <Text style={styles.helper}>
                Public address: https://api.luxuryimportsusa.shop
              </Text>
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
                  {detailRows.map(([label, value]) => (
                    <FieldRow key={label} label={label} value={value} />
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
                  <Text style={styles.helper}>Save photos straight into this claim folder on your PC.</Text>
                  <View style={styles.photoRow}>
                    <Action
                      label={uploadingPhoto ? "Uploading..." : "Take Photo"}
                      onPress={() => uploadPhoto(true)}
                      compact
                    />
                    <Action label="Upload Photos" onPress={() => uploadPhoto(false)} compact />
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
                  {(claimDetail.files || []).length ? (
                    claimDetail.files.map((file) => (
                      <Pressable key={file.relative_path} style={styles.listItem} onPress={() => Linking.openURL(file.file_url)}>
                        <Text style={styles.itemTitle}>{file.name}</Text>
                        <Text style={styles.itemMeta}>{file.relative_path}</Text>
                      </Pressable>
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
    </SafeAreaView>
  );

  function renderContent() {
    if (homeTab === "claims") {
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

function FieldRow({ label, value }) {
  return (
    <View style={styles.fieldRow}>
      <Text style={styles.fieldLabel}>{label}</Text>
      <Text style={styles.fieldValue}>{value || "-"}</Text>
    </View>
  );
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

function RouteRow({ item, onPress, badge, actions = [], editor = null }) {
  return (
    <View style={styles.listItem}>
      {badge ? <Text style={styles.routeBadge}>{badge}</Text> : null}
      <Pressable onPress={onPress}>
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
    paddingHorizontal: 14,
    paddingTop: Platform.OS === "android" ? (NativeStatusBar.currentHeight || 0) + 8 : 6,
    paddingBottom: 0,
    gap: 6,
  },
  headerTop: {
    flexDirection: "row",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 10,
  },
  header: {
    flex: 1,
    gap: 0,
  },
  eyebrow: {
    fontSize: 10,
    color: "#6a7f9d",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: 1.2,
  },
  title: {
    fontSize: 17,
    fontWeight: "800",
    color: "#0f2742",
  },
  subtitle: {
    fontSize: 11,
    color: "#5d718d",
    lineHeight: 16,
  },
  versionText: {
    marginTop: 6,
    fontSize: 12,
    fontWeight: "800",
    color: "#b42318",
  },
  versionSubtext: {
    fontSize: 10,
    color: "#6a7f9d",
    fontWeight: "700",
  },
  card: {
    backgroundColor: "#ffffff",
    borderRadius: 20,
    padding: 16,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    gap: 10,
  },
  utilityCard: {
    backgroundColor: "#ffffff",
    borderRadius: 18,
    padding: 12,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    gap: 8,
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
    gap: 6,
    backgroundColor: "#ffffff",
    borderRadius: 999,
    paddingHorizontal: 10,
    paddingVertical: 5,
    borderWidth: 1,
    borderColor: "#d7e1ee",
    maxWidth: 210,
  },
  contentWrap: {
    flex: 1,
  },
  statusChipLabel: {
    fontSize: 10,
    fontWeight: "800",
    color: "#6b809d",
    textTransform: "uppercase",
    letterSpacing: 0.8,
  },
  statusChipValue: {
    flex: 1,
    fontSize: 11,
    color: "#183659",
    fontWeight: "600",
  },
  statusMessage: {
    flex: 1,
    fontSize: 11,
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
    fontSize: 13,
    color: "#6e84a3",
    lineHeight: 18,
  },
  helperStrong: {
    fontSize: 12,
    fontWeight: "800",
    color: "#183659",
  },
  input: {
    borderWidth: 1,
    borderColor: "#cad7e8",
    backgroundColor: "#fbfdff",
    borderRadius: 14,
    paddingHorizontal: 14,
    paddingVertical: 11,
    fontSize: 15,
    color: "#10253f",
  },
  search: {
    borderWidth: 1,
    borderColor: "#cad7e8",
    backgroundColor: "#ffffff",
    borderRadius: 15,
    paddingHorizontal: 15,
    paddingVertical: 9,
    fontSize: 13,
    color: "#10253f",
  },
  tabRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 5,
  },
  chip: {
    minWidth: 68,
    paddingHorizontal: 11,
    paddingVertical: 6,
    borderRadius: 999,
    backgroundColor: "#e7eef7",
    alignItems: "center",
  },
  chipOn: {
    backgroundColor: "#103c6d",
  },
  chipText: {
    color: "#35506f",
    fontSize: 11,
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
    fontSize: 12,
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
    fontSize: 14,
    lineHeight: 20,
  },
  listItem: {
    marginHorizontal: 10,
    marginVertical: 4,
    paddingHorizontal: 13,
    paddingVertical: 10,
    backgroundColor: "#ffffff",
    borderRadius: 16,
    borderWidth: 1,
    borderColor: "#e1e9f3",
  },
  itemTitle: {
    fontSize: 14,
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
    marginTop: 4,
    fontSize: 11,
    lineHeight: 16,
    color: "#5d718d",
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
    fontSize: 20,
    fontWeight: "800",
    color: "#10253f",
  },
  smallBtn: {
    backgroundColor: "#e8eef7",
    borderRadius: 14,
    paddingHorizontal: 14,
    paddingVertical: 10,
  },
  smallBtnText: {
    color: "#14385f",
    fontWeight: "700",
    fontSize: 14,
  },
  scrollPad: {
    paddingBottom: 24,
    gap: 12,
  },
  section: {
    fontSize: 20,
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
    fontSize: 13,
    color: "#6e84a3",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: 0.8,
  },
  fieldValue: {
    marginTop: 4,
    fontSize: 18,
    color: "#10253f",
    lineHeight: 24,
  },
  stack: {
    gap: 10,
  },
  action: {
    backgroundColor: "#e8eef7",
    borderRadius: 16,
    paddingVertical: 13,
    paddingHorizontal: 16,
    alignItems: "center",
  },
  actionCompact: {
    flex: 1,
    paddingVertical: 11,
    paddingHorizontal: 12,
  },
  actionPrimary: {
    backgroundColor: "#103c6d",
  },
  actionDisabled: {
    opacity: 0.55,
  },
  actionText: {
    fontSize: 16,
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
