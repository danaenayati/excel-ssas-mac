// Load-once data for the Explorer (databases, cubes, fields), cached by key for as
// long as the pane is open so that switching between reports does not ask again.

import { useEffect, useRef, useState } from "react";
import { toUiError, type UiError } from "./runner";

export interface Loadable<T> {
  data: T | null;
  error: UiError | null;
  loading: boolean;
  retry: () => void;
}

interface Slot<T> {
  key: string | null;
  data: T | null;
  error: UiError | null;
  loading: boolean;
}

export type Cache = Map<string, Promise<unknown>>;

/**
 * `key` identifies what to load; null means "nothing to load yet". A failed load is
 * not cached, so `retry()` (or simply asking again later) tries afresh.
 */
export function useCached<T>(cache: Cache, key: string | null, load: () => Promise<T>): Loadable<T> {
  const [slot, setSlot] = useState<Slot<T>>({ key: null, data: null, error: null, loading: false });
  const [attempt, setAttempt] = useState(0);
  const loader = useRef(load);
  loader.current = load;

  useEffect(() => {
    if (key === null) {
      setSlot({ key: null, data: null, error: null, loading: false });
      return;
    }
    let alive = true;
    let promise = cache.get(key) as Promise<T> | undefined;
    if (!promise) {
      const fresh = loader.current();
      promise = fresh;
      cache.set(key, fresh);
      fresh.catch(() => {
        if (cache.get(key) === fresh) cache.delete(key);
      });
    }
    setSlot({ key, data: null, error: null, loading: true });
    promise.then(
      (data) => {
        if (alive) setSlot({ key, data, error: null, loading: false });
      },
      (failure) => {
        if (alive) setSlot({ key, data: null, error: toUiError(failure), loading: false });
      },
    );
    return () => {
      alive = false;
    };
  }, [cache, key, attempt]);

  // On the render right after `key` changes the slot still describes the old key.
  const current = slot.key === key;
  return {
    data: current ? slot.data : null,
    error: current ? slot.error : null,
    loading: key !== null && (!current || slot.loading),
    retry: () => setAttempt((n) => n + 1),
  };
}
