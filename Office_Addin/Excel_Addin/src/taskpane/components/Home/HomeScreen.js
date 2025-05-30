import React, { useState, useEffect } from "react";
import jwtDecode from "jwt-decode"; // note: import default, not named
import "./HomeScreen.scss";
import logo from "../assets/images/logoword.png";
import logError from "../assets/images/LogError.png";
import dismiss from "../assets/images/Dismiss.png";
import { useNavigate } from "react-router-dom";

const checkEmail = async (email) => {
  try {
    const response = await fetch(
      `https://us-central1-bbca-be.cloudfunctions.net/api/check-email?email=${email}`
    );
    const data = await response.json();
    return data.exists ?? false;
  } catch (error) {
    console.error("API call failed:", error);
    return false;
  }
};

const HomeScreen = () => {
  const navigate = useNavigate();
  const [officeReady, setOfficeReady] = useState(false);
  const [loading, setLoading] = useState(false);
  const [showError, setShowError] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!window.Office) {
      setError("Office.js not loaded - check manifest or deployment path.");
      return;
    }

    Office.onReady(() => {
      console.log("Office is ready");
      setOfficeReady(true);
    });
  }, []);

 

  const handleLogin = () => {
    if (!officeReady) {
      setError("Office is not ready yet.");
      return;
    }

    setLoading(true);

    Office.context.auth.getAccessTokenAsync(
      {
        allowConsentPrompt: true,
        allowSignInPrompt: true,
        forMSGraphAccess: true,
      },
      async (result) => {
        if (result.status === "succeeded" && result.value) {
          try {
            const decodedToken = jwtDecode(result.value);
            const email = decodedToken.preferred_username || decodedToken.upn || decodedToken.email;

            console.log("Decoded Email:", email);

            if (!email) {
              setLoading(false);
              setError("Could not extract email from token.");
              return;
            }

            const emailCheck = await checkEmail(email);

            setLoading(false);

            if (emailCheck) {
              setShowError(false);
              navigate("/exportExcel");
            } else {
              setShowError(true);
            }
          } catch (decodeError) {
            setLoading(false);
            setError("Failed to decode token.");
          }
        } else {
          setLoading(false);
          const errorMsg = result.error?.message || "Unknown error during token retrieval.";
          console.error("Token retrieval failed:", errorMsg);
          setError(`Token retrieval failed: ${errorMsg}`);
        }
      }
    );
  };

  if (error) {
    return (
      <div style={{ color: "red", padding: "20px" }}>
        CRITICAL ERROR: {error}
        <br />
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
                    This account is not in our system. Please contact the administrator.
                  </p>
                </div>
                <button className="dismiss-button" onClick={() => setShowError(false)}>
                  <img src={dismiss} alt="dismiss icon" />
                </button>
              </div>
            </div>
          )}
          <div className="content">
            <img src={logo} alt="logo" className="logo" />
            <button className="login-button" onClick={handleLogin}>
              Sign In With Microsoft
            </button>
          </div>
        </>
      )}
    </div>
  );
};

export default HomeScreen;
