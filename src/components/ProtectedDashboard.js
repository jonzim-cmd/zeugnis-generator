// src/components/ProtectedDashboard.js
import React, { useState, useEffect } from 'react';
import Dashboard from './Dashboard';

const ProtectedDashboard = (props) => {
  const [authenticated, setAuthenticated] = useState(false);
  const [password, setPassword] = useState('');
  const [attempts, setAttempts] = useState(0);
  const [lockoutUntil, setLockoutUntil] = useState(null);
  const [showPasswordPrompt, setShowPasswordPrompt] = useState(true);

  // Beim Laden der Komponente: Versuche, gespeicherte Sperrzeit und Fehlversuche aus localStorage zu laden
  useEffect(() => {
    const storedLockout = localStorage.getItem('lockoutUntil');
    if (storedLockout) {
      const lockoutTime = parseInt(storedLockout, 10);
      if (Date.now() < lockoutTime) {
        setLockoutUntil(lockoutTime);
      } else {
        localStorage.removeItem('lockoutUntil');
      }
    }
    const storedAttempts = localStorage.getItem('attempts');
    if (storedAttempts) {
      setAttempts(parseInt(storedAttempts, 10));
    }
  }, []);

  // Aktualisiere localStorage, wenn sich lockoutUntil ändert
  useEffect(() => {
    if (lockoutUntil) {
      localStorage.setItem('lockoutUntil', lockoutUntil);
    } else {
      localStorage.removeItem('lockoutUntil');
    }
  }, [lockoutUntil]);

  // Speichere die Fehlversuche in localStorage
  useEffect(() => {
    localStorage.setItem('attempts', attempts);
  }, [attempts]);

  const handlePasswordSubmit = async (e) => {
    e.preventDefault();
    // Falls noch eine Sperrzeit aktiv ist, abbrechen
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
        // Bei korrektem Passwort: Authentifizierung erfolgreich und lokale Daten löschen
        setAuthenticated(true);
        setShowPasswordPrompt(false);
        localStorage.removeItem('lockoutUntil');
        localStorage.removeItem('attempts');
      } else {
        // Bei falschem Passwort: Fehlversuche erhöhen und ggf. Sperrzeit setzen
        const newAttempts = attempts + 1;
        setAttempts(newAttempts);
        if (newAttempts >= 5) {
          const newLockout = Date.now() + 60 * 60 * 1000; // 60 Minuten Sperrzeit
          setLockoutUntil(newLockout);
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
        <form 
          onSubmit={handlePasswordSubmit} 
          style={{ border: '1px solid #ccc', padding: '1rem', display: 'inline-block' }}
        >
          <label>Bitte gib das Passwort ein:</label>
          <br />
          <input 
            type="password" 
            value={password} 
            onChange={(e) => setPassword(e.target.value)} 
            style={{ marginTop: '0.5rem' }}
          />
          <br />
          <button type="submit" style={{ marginTop: '0.5rem' }}>
            Absenden
          </button>
          {attempts > 0 && <p>Fehlversuche: {attempts} von 5</p>}
          {lockoutUntil && Date.now() < lockoutUntil && (
            <p style={{ color: 'red' }}>
              Sperre aktiv. Bitte warte bis {new Date(lockoutUntil).toLocaleTimeString()}.
            </p>
          )}
        </form>
      </div>
    );
  }

  // Bei erfolgreicher Authentifizierung wird das Dashboard gerendert
  return <Dashboard {...props} />;
};

export default ProtectedDashboard;
