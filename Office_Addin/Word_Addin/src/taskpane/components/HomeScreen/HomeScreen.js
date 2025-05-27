import React, { useState, useEffect } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import jwtDecode from "jwt-decode";  // import without braces
import { useNavigate } from "react-router-dom";

import "./HomeScreen.scss";
import logo from "../../../../assets/logoword.png";
import logError from "../../../../assets/LogError.png";
import dismiss from "../../../../assets/Dismiss.png";
import microsoftTri from "../../../../assets/MicrosoftTri.png";
import needHelp from "../../../../assets/needHelp.png";

const CLIENT_ID = "ab1349c6-78b8-4824-800b-066ea1c49997";
const AUTHORITY = "https://login.microsoftonline.com/common";

const checkEmail = async (email) => {
  try {
    const res = await fetch(
      `https://us-central1-bbca-be.cloudfunctions.net/api/check-email?email=${email}`
    );
    const data = await res.json();
    return data.exists === true;
  } catch (err) {
    console.error("API call failed:", err);
    return false;
  }
};

const HomeScreen = () => {
  const navigate = useNavigate();
  const [officeReady, setOfficeReady] = useState(false);
  const [error, setError] = useState("");
  const [showError, setShowError] = useState(false);
  const [loading, setLoading] = useState(false);
  const [email, setEmail] = useState("");

  // Initialize MSAL instance once
  const msalInstance = new PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: AUTHORITY,
      redirectUri: window.location.origin,
    },
  });

  // Office.js initialization
  useEffect(() => {
    if (window.Office) {
      window.Office.onReady(() => {
        setOfficeReady(true);
      });
    } else {
      setError("Office.js not loaded - check manifest URLs");
    }
  }, []);

  // Initialize MSAL and login on component mount
  useEffect(() => {
    const initAndLogin = async () => {
      try {
        await msalInstance.initialize();
        const response = await msalInstance.loginPopup({
          scopes: ["openid", "email", "profile"],
        });
        const decoded = jwtDecode(response.idToken);
        setEmail(decoded.email || decoded.upn || "");
      } catch (err) {
        console.error("Login failed:", err);
        setError("Login failed. Please try again.");
      }
    };

    initAndLogin();
  }, []); // run once

  // Handle Office.js auth button click
  const handleLogin = () => {
    if (!officeReady) return;
    setLoading(true);

    window.Office.context.auth.getAccessTokenAsync(
      {
        allowConsentPrompt: true,
        allowSignInPrompt: true,
        forMSGraphAccess: true,
      },
      async (result) => {
        if (result.status === "succeeded" && result.value) {
          try {
            const decoded = jwtDecode(result.value);
            const userEmail = decoded.preferred_username || decoded.email;
            setEmail(userEmail);

            const exists = await checkEmail(userEmail);
            setLoading(false);

            if (exists) {
              setShowError(false);
              navigate(`/Home/${userEmail}`);
            } else {
              setShowError(true);
            }
          } catch (e) {
            console.error("Token decoding or API check failed:", e);
            setError("Authentication failed, please try again.");
            setLoading(false);
          }
        } else {
          setError("Token retrieval failed.");
          setLoading(false);
        }
      }
    );
  };

  if (error) {
    return (
      <div style={{ color: "red", padding: "20px" }}>
        CRITICAL ERROR: {error}
        <button onClick={() => window.location.reload()}>Reload Add-in</button>
      </div>
    );
  }

  return (
    <div className="container">
      {loading ? (
        <div className="loading-container">
          <div className="spinner"></div>
          <p className="loading-text">Loading...</p>
        </div>
      ) : (
        <>
          {showError && (
            <div className="error-container">
              <div className="error-content">
                <img src={logError} alt="error icon" className="error-icon" />
                <div className="error-text-container">
                  <p className="error-title">Log In Error</p>
                  <p className="error-message">
                    This account is not in our system. For questions contact the administrator.
                  </p>
                </div>
                <button className="dismiss-button" onClick={() => setShowError(false)}>
                  <img src={dismiss} alt="dismiss icon" className="dismiss-icon" />
                </button>
              </div>
            </div>
          )}

          <div className="main-content">
            <img src={logo} alt="logo" className="logo" />
            <h1 className="welcome-text">Welcome!</h1>
            <p className="sub-text">Export from Excel to Word with ease.</p>
            <button className="login-button" onClick={handleLogin}>
              <img src={microsoftTri} alt="Microsoft icon" className="microsoft-icon" />
              <span>Sign In With Microsoft</span>
            </button>
          </div>

          <div className="help-container">
            <img src={needHelp} alt="help icon" className="help-icon" />
            <div>
              <p className="help-title">Need Help?</p>
              <p className="help-text">Contact Administrator</p>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default HomeScreen;
