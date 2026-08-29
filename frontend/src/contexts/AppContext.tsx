import {
  createContext,
  useContext,
  useReducer,
  useEffect,
  useMemo
} from 'react';

import type {
  ReactNode,
  Dispatch
} from 'react';

import { useMsal } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';

import type {
  AppState,
  AppAction
} from '../types/appState';

import { initialAppState } from '../types/appState';
import { appReducer } from '../reducers/appReducer';

interface AppContextValue {
  state: AppState;
  dispatch: Dispatch<AppAction>;
}

const AppContext = createContext<AppContextValue | undefined>(undefined);

// Lightweight dev logger prevents accidental prod noise
const devLogger = {
  enabled: import.meta.env.DEV,

  group(label: string) {
    if (this.enabled) {
      console.group(label);
    }
  },

  log: function (...args: unknown[]) {
    if (this.enabled) {
      console.log(...args);
    }
  },

  end() {
    if (this.enabled) {
      console.groupEnd();
    }
  }
};

// Dev mode logging middleware (diff-based)
const logStateChange = (
  action: AppAction,
  prevState: AppState,
  nextState: AppState
) => {
  if (!devLogger.enabled) {
    return;
  }

  const timestamp = new Date()
    .toISOString()
    .split('T')[1]
    .split('.')[0];

  devLogger.group(`🔄 [${timestamp}] ${action.type}`);
  devLogger.log('Action:', action);

  const changes: Record<string, unknown> = {};

  // Track all meaningful state changes
  if (prevState.auth.status !== nextState.auth.status) {
    changes['auth.status'] =
      `${prevState.auth.status} → ${nextState.auth.status}`;
  }

  if (prevState.chat.status !== nextState.chat.status) {
    changes['chat.status'] =
      `${prevState.chat.status} → ${nextState.chat.status}`;
  }

  if (
    prevState.chat.messages.length !==
    nextState.chat.messages.length
  ) {
    changes['chat.messages.length'] =
      `${prevState.chat.messages.length} → ${nextState.chat.messages.length}`;
  }

  if (
    prevState.chat.streamingMessageId !==
    nextState.chat.streamingMessageId
  ) {
    changes['chat.streamingMessageId'] =
      `${prevState.chat.streamingMessageId} → ${nextState.chat.streamingMessageId}`;
  }

  if (
    prevState.ui.chatInputEnabled !==
    nextState.ui.chatInputEnabled
  ) {
    changes['ui.chatInputEnabled'] =
      `${prevState.ui.chatInputEnabled} → ${nextState.ui.chatInputEnabled}`;
  }

  if (
    prevState.conversations.sidebarOpen !==
    nextState.conversations.sidebarOpen
  ) {
    changes['conversations.sidebarOpen'] =
      `${prevState.conversations.sidebarOpen} → ${nextState.conversations.sidebarOpen}`;
  }

  if (
    prevState.conversations.list.length !==
    nextState.conversations.list.length
  ) {
    changes['conversations.list.length'] =
      `${prevState.conversations.list.length} → ${nextState.conversations.list.length}`;
  }

  if (Object.keys(changes).length) {
    devLogger.log('Changes:', changes);
  } else {
    devLogger.log('(No state changes)');
  }

  devLogger.end();
};

/**
 * Enhanced reducer with logging middleware
 */
const reducerWithLogging = (
  state: AppState,
  action: AppAction
): AppState => {
  const nextState = appReducer(state, action);

  logStateChange(
    action,
    state,
    nextState
  );

  return nextState;
};

export const AppProvider: React.FC<{
  children: ReactNode;
}> = ({ children }) => {

  const [state, dispatch] = useReducer(
    reducerWithLogging,
    initialAppState
  );

  const {
    accounts,
    inProgress
  } = useMsal();

  /**
   * Synchronize application authentication state
   * with Microsoft Authentication Library (MSAL).
   *
   * IMPORTANT:
   * Wait until MSAL finishes processing redirects
   * before deciding whether the user is authenticated.
   */
  useEffect(() => {

    // MSAL is still processing login/redirect.
    // Keep application in initializing state.
    if (inProgress !== InteractionStatus.None) {
      return;
    }

    // User exists in MSAL
    if (accounts.length > 0) {

      const account = accounts[0];

      // Avoid unnecessary dispatches
      if (
        state.auth.status !== 'authenticated' ||
        state.auth.user?.homeAccountId !== account.homeAccountId
      ) {
        dispatch({
          type: 'AUTH_INITIALIZED',
          user: account
        });
      }

      return;
    }

    /**
     * MSAL has finished initialization,
     * but no signed-in account exists.
     *
     * The previous implementation did nothing here,
     * which allowed auth.status to remain
     * "initializing" forever.
     */
    if (state.auth.status !== 'unauthenticated') {
      dispatch({
        type: 'AUTH_TOKEN_EXPIRED'
      });
    }

  }, [
    accounts,
    inProgress,
    state.auth.status,
    state.auth.user
  ]);

  // Dev mode: Log when provider mounts and unmounts
  useEffect(() => {

    devLogger.log('🚀 AppProvider initialized');

    return () => {
      devLogger.log('🔌 AppProvider unmounted');
    };

  }, []);

  // Memoize context value to prevent unnecessary re-renders
  const contextValue = useMemo(
    () => ({
      state,
      dispatch
    }),
    [
      state,
      dispatch
    ]
  );

  return (
    <AppContext.Provider value={contextValue}>
      {children}
    </AppContext.Provider>
  );
};

/**
 * Hook to access app state and dispatch.
 * Throws error if used outside AppProvider.
 */
// eslint-disable-next-line react-refresh/only-export-components
export const useAppContext = () => {

  const context = useContext(AppContext);

  if (!context) {
    throw new Error(
      'useAppContext must be used within AppProvider'
    );
  }

  return context;
};
