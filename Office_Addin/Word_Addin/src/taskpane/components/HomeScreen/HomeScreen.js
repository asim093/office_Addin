import React, { useState, useEffect, useRef } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import jwtDecode from "jwt-decode";
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
  const msalInstanceRef = useRef(null);
  const [msalInitialized, setMsalInitialized] = useState(false);

  // Initialize MSAL instance only once
  useEffect(() => {
    const initializeMsal = async () => {
      try {
        if (!msalInstanceRef.current) {
          msalInstanceRef.current = new PublicClientApplication({
            auth: {
              clientId: CLIENT_ID,
              authority: AUTHORITY,
              redirectUri: window.location.origin,
            },
            cache: {
              cacheLocation: "sessionStorage", // Production mein better performance
              storeAuthStateInCookie: false,
            }
          });

          // Initialize MSAL properly
          await msalInstanceRef.current.initialize();
          setMsalInitialized(true);
          console.log("MSAL initialized successfully");
        }
      } catch (err) {
        console.error("MSAL initialization failed:", err);
        setError("Authentication system failed to initialize");
      }
    };

    initializeMsal();
  }, []);

  // Office.js initialization
  useEffect(() => {
    if (window.Office) {
      window.Office.onReady(() => {
        console.log("Office.js ready");
        setOfficeReady(true);
      });
    } else {
      console.log("Office.js not available");
      setOfficeReady(true); // For testing outside Office environment
    }
  }, []);

  // Auto-login after MSAL is initialized
  useEffect(() => {
    const performAutoLogin = async () => {
      if (!msalInitialized || !msalInstanceRef.current) return;

      try {
        // Check if user is already logged in
        const accounts = msalInstanceRef.current.getAllAccounts();
        if (accounts.length > 0) {
          const account = accounts[0];
          setEmail(account.username);
          console.log("User already logged in:", account.username);
          return;
        }

        // Try silent login first
        const silentRequest = {
          scopes: ["openid", "email", "profile"],
          account: accounts[0] || null,
        };

        const response = await msalInstanceRef.current.acquireTokenSilent(silentRequest);
        if (response && response.idToken) {
          const decoded = jwtDecode(response.idToken);
          setEmail(decoded.email || decoded.upn || decoded.preferred_username || "");
          console.log("Silent login successful");
        }
      } catch (err) {
        console.log("Silent login failed, user needs to login manually:", err.message);
        // This is normal - user just needs to click login button
      }
    };

    performAutoLogin();
  }, [msalInitialized]);

  // Handle login button click with better error handling
  const handleMsalLogin = async () => {
    if (!msalInitialized || !msalInstanceRef.current) {
      setError("Authentication system not ready. Please wait and try again.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const loginRequest = {
        scopes: ["openid", "email", "profile"],
        prompt: "select_account"
      };

      const response = await msalInstanceRef.current.loginPopup(loginRequest);
      
      if (response && response.idToken) {
        const decoded = jwtDecode(response.idToken);
        const userEmail = decoded.email || decoded.upn || decoded.preferred_username;
        
        if (!userEmail) {
          throw new Error("No email found in authentication response");
        }

        setEmail(userEmail);
        console.log("Login successful:", userEmail);

        // Check if user exists in system
        const exists = await checkEmail(userEmail);
        
        if (exists) {
          setShowError(false);
          navigate(`/Home/${userEmail}`);
        } else {
          setShowError(true);
        }
      } else {
        throw new Error("No valid response received from authentication");
      }
    } catch (err) {
      console.error("MSAL login failed:", err);
      setError(`Login failed: ${err.message || "Please try again"}`);
    } finally {
      setLoading(false);
    }
  };

  // Handle Office.js auth (fallback method)
  const handleOfficeLogin = () => {
    if (!window.Office || !window.Office.context) {
      handleMsalLogin(); // Fallback to MSAL if Office.js not available
      return;
    }

    setLoading(true);
    setError("");

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
            const userEmail = decoded.preferred_username || decoded.email || decoded.upn;
            
            if (!userEmail) {
              throw new Error("No email found in token");
            }

            setEmail(userEmail);
            console.log("Office login successful:", userEmail);

            const exists = await checkEmail(userEmail);
            
            if (exists) {
              setShowError(false);
              navigate(`/Home/${userEmail}`);
            } else {
              setShowError(true);
            }
          } catch (e) {
            console.error("Office token processing failed:", e);
            setError("Authentication failed, please try again.");
          }
        } else {
          console.error("Office auth failed:", result);
          setError("Office authentication failed. Trying alternative method...");
          // Fallback to MSAL
          setTimeout(() => handleMsalLogin(), 1000);
        }
        setLoading(false);
      }
    );
  };

  // Main login handler - tries Office.js first, then MSAL
  const handleLogin = () => {
    if (officeReady && window.Office && window.Office.context && window.Office.context.auth) {
      handleOfficeLogin();
    } else {
      handleMsalLogin();
    }
  };

  // Error display component
  if (error && !showError) {
    return (
      <div style={{ 
        color: "red", 
        padding: "20px", 
        textAlign: "center",
        backgroundColor: "#ffebee",
        border: "1px solid #f44336",
        borderRadius: "4px",
        margin: "20px"
      }}>
        <h3>Authentication Error</h3>
        <p>{error}</p>
        <button 
          onClick={() => {
            setError("");
            window.location.reload();
          }}
          style={{
            padding: "10px 20px",
            backgroundColor: "#f44336",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            marginTop: "10px"
          }}
        >
          Reload Add-in
        </button>
      </div>
    );
  }

  return (
    <div className="container">
      {loading ? (
        <div className="loading-container">
          <div className="spinner"></div>
          <p className="loading-text">Authenticating...</p>
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
                    This account ({email}) is not in our system. For questions contact the administrator.
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
            <button 
              className="login-button" 
              onClick={handleLogin}
              disabled={loading || (!msalInitialized && !officeReady)}
            >
              <img src={microsoftTri} alt="Microsoft icon" className="microsoft-icon" />
              <span>
                {!msalInitialized && !officeReady ? "Loading..." : "Sign In With Microsoft"}
              </span>
            </button>
            {email && (
              <p style={{ marginTop: "10px", fontSize: "12px", color: "#666" }}>
                Current user: {email}
              </p>
            )}
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