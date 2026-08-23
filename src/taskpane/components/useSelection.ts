/**
 * Keeps the pane in step with what the user has selected in the grid, so each
 * command can show the exact range it is about to touch before it touches it.
 */

import * as React from "react";

import { CleaningScope, describeTarget, readTargetHeaders } from "../features/dataCleaning";

/** Subscribe to Excel's selection-changed event; returns an unsubscribe function. */
function subscribeToSelection(onChange: () => void): () => void {
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
  } catch {
    // Older hosts may not raise this event; the pane still works, just without
    // live updates to the "acting on" hint.
    return () => undefined;
  }

  return () => {
    try {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
        handler: onChange,
      });
    } catch {
      // Nothing useful to do if the host refuses to detach the handler.
    }
  };
}

/** The address a command with this scope would act on, kept current as the user clicks around. */
export function useTargetAddress(scope: CleaningScope, refreshToken = 0): string {
  const [address, setAddress] = React.useState("...");

  React.useEffect(() => {
    let active = true;

    const update = () => {
      void describeTarget(scope).then((next) => {
        if (active) {
          setAddress(next);
        }
      });
    };

    update();
    const unsubscribe = subscribeToSelection(update);

    return () => {
      active = false;
      unsubscribe();
    };
  }, [scope, refreshToken]);

  return address;
}

/** Header labels of the current target, for the key-column picker. */
export function useTargetHeaders(scope: CleaningScope, enabled: boolean): string[] {
  const [headers, setHeaders] = React.useState<string[]>([]);

  React.useEffect(() => {
    if (!enabled) {
      return undefined;
    }
    let active = true;

    const update = () => {
      void readTargetHeaders(scope).then((next) => {
        if (active) {
          setHeaders(next);
        }
      });
    };

    update();
    const unsubscribe = subscribeToSelection(update);

    return () => {
      active = false;
      unsubscribe();
    };
  }, [scope, enabled]);

  return headers;
}
