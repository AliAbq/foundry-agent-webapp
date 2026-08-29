import { useMsal } from "@azure/msal-react";
import {
  InteractionRequiredAuthError,
  InteractionStatus,
} from "@azure/msal-browser";

import {
  tokenRequest,
  loginRequest,
} from "../config/authConfig";

import {
  useCallback,
  useMemo,
} from "react";

/**
 * Authentication hook for MSAL-based authentication.
 *
 * - Attempts silent token renewal first.
 * - Redirects to Microsoft sign-in when interaction is required.
 * - Provides a manual signInAgain function if needed by the UI.
 */
export const useAuth = () => {
  const {
    instance,
    accounts,
    inProgress,
  } = useMsal();

  /**
   * Manually restart Microsoft authentication.
   */
  const signInAgain = useCallback(async () => {
    if (inProgress !== InteractionStatus.None) {
      return;
    }

    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error(
        "Failed to restart sign-in:",
        error
      );
    }
  }, [
    instance,
    inProgress,
  ]);

  /**
   * Acquire an access token.
   *
   * Silent renewal is attempted first.
   * If Microsoft requires user interaction,
   * redirect the browser to Microsoft Entra login.
   */
  const getAccessToken = useCallback(
    async (): Promise<string | null> => {

      /**
       * No account exists.
       * Start sign-in again once MSAL is idle.
       */
      if (accounts.length === 0) {
        if (inProgress === InteractionStatus.None) {
          try {
            await instance.loginRedirect(loginRequest);
          } catch (error) {
            console.error(
              "Login redirect failed:",
              error
            );
          }
        }

        return null;
      }

      const account =
        instance.getActiveAccount() ||
        accounts[0];

      const request = {
        ...tokenRequest,
        account,
      };

      try {
        /**
         * First attempt:
         * silently obtain/renew the token.
         */
        const response =
          await instance.acquireTokenSilent(
            request
          );

        return response.accessToken;

      } catch (error) {

        /**
         * Microsoft requires user interaction.
         *
         * Instead of returning null and leaving
         * the application on an expired-session
         * screen, redirect back to Microsoft login.
         */
        if (
          error instanceof
          InteractionRequiredAuthError
        ) {
          console.warn(
            "Silent token acquisition requires interaction. Redirecting to sign in."
          );

          if (
            inProgress ===
            InteractionStatus.None
          ) {
            try {
              await instance.acquireTokenRedirect(
                request
              );
            } catch (redirectError) {
              console.error(
                "Token redirect failed:",
                redirectError
              );
            }
          }

          return null;
        }

        console.error(
          "Token acquisition error:",
          error
        );

        return null;
      }
    },
    [
      instance,
      accounts,
      inProgress,
    ]
  );

  /**
   * Authentication status
   */
  const isAuthenticated = useMemo(
    () => accounts.length > 0,
    [accounts.length]
  );

  /**
   * Current user
   */
  const user = useMemo(
    () =>
      instance.getActiveAccount() ||
      accounts[0] ||
      null,
    [
      instance,
      accounts,
    ]
  );

  return useMemo(
    () => ({
      getAccessToken,
      signInAgain,
      isAuthenticated,
      user,
    }),
    [
      getAccessToken,
      signInAgain,
      isAuthenticated,
      user,
    ]
  );
};
