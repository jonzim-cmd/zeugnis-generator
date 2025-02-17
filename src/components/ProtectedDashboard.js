// src/components/ProtectedDashboard.js
import React, { useState } from 'react';
import Dashboard from './Dashboard';

const ProtectedDashboard = (props) => {
  const [authenticated, setAuthenticated] = useState(false);
  const [password, setPassword] = useState('');
  const [attempts, setAttempts] = useState(0);
  const [lockoutUntil, setLockoutUntil] = useState(null);
  const [showPasswordPrompt, setShowPasswordPrompt] = useState(true);

  const handlePasswordSubmit = async () => {
    if (lockoutUntil && Date.now() < lockoutUntil) {
      alert("Zu viele Fehlversuche. Bitte warte, bis die Sperrzeit abgelaufen ist.");
      return;
    }
    try {
      const res = await fetch('/api/validate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password })
      });
      if (res.ok) {
        setAuthenticated(true);
        setShowPasswordPrompt(false);
      } else {
        const newAttempts = attempts + 1;
        setAttempts(newAttempts);
        if (newAttempts >= 5) {
          setLockoutUntil(Date.now() + 60 * 60 * 1000); // 60 Minuten Sperrzeit
          alert("Zu viele Fehlversuche. Du bist für 60 Minuten gesperrt.");
        } else {
          alert(`Falsches Passwort. Versuch ${newAttempts} von 5.`);
        }
      }
    } catch (error) {
      console.error("Fehler bei der Passwortüberprüfung:", error);
      alert("Fehler bei der Passwortüberprüfung.");
    }
  };

  if (!authenticated) {
    return (
      <div style={{ textAlign: 'center', marginTop: '2rem' }}>
        <div style={{ border: '1px solid #ccc', padding: '1rem', display: 'inline-block' }}>
          <label>Bitte gib das Passwort ein:</label>
          <br />
          <input 
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            style={{ marginTop: '0.5rem' }}
          />
          <br />
          <button onClick={handlePasswordSubmit} style={{ marginTop: '0.5rem' }}>
            Absenden
          </button>
          {attempts > 0 && <p>Fehlversuche: {attempts} von 5</p>}
          {lockoutUntil && Date.now() < lockoutUntil && (
            <p style={{ color: 'red' }}>
              Sperre aktiv. Bitte warte bis {new Date(lockoutUntil).toLocaleTimeString()}.
            </p>
          )}
        </div>
      </div>
    );
  }

  // Bei erfolgreicher Authentifizierung wird das Dashboard gerendert
  return <Dashboard {...props} />;
};

export default ProtectedDashboard;
