import React, { useState, useEffect } from "react";
import { jwtDecode } from "jwt-decode";
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
    if (!officeReady) return;

    setLoading(true);

    Office.context.auth.getAccessTokenAsync(
      {
        allowConsentPrompt: true,
        allowSignInPrompt: true,
        forMSGraphAccess: true,
      },
      async (result) => {
        console.log("Token callback result:", result);

        if (result.status === "succeeded" && result.value) {
          const decodedToken = jwtDecode(result.value);
          const email = decodedToken.preferred_username;

          console.log("Email:", email);

          const emailCheck = await checkEmail(email);
          console.log("Email Check:", emailCheck);

          setLoading(false);
          if (emailCheck) {
            setShowError(false);
            navigate("/exportExcel");
          } else {
            setShowError(true);
          }
        } else {
          setLoading(false);
          console.error("Token retrieval failed:", result.error);
          setError(`Token retrieval failed: ${result.error?.message || 'Unknown error'}`);
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