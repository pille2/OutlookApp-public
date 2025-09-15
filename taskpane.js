// CRM Email Manager - Taskpane JavaScript
Office.onReady((info) => {
    console.log("Office.onReady called", info);
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook host detected");
        if (document.readyState === 'loading') {
            document.addEventListener("DOMContentLoaded", initializeApp);
        } else {
            initializeApp();
        }
    } else {
        console.error("Wrong host type:", info.host);
    }
});

let currentEmail = null;
let emailData = null;
let debugLogs = [];
let errorLogs = [];

async function initializeApp() {
    console.log("ALL ONE Lead Tracker v2.17.3 initialisiert");
    addDebugLog("App initialisiert");
    
    // Debug Panel Setup
    setupDebugPanel();
    
    // Event Listener für Rating Buttons
    document.querySelectorAll('.rating-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            selectRating(this.dataset.rating);
        });
    });
    
    // Event Listener für Action Buttons
    document.getElementById('sendBtn').addEventListener('click', sendToCRM);
    
    // Lade E-Mail Informationen
    await loadEmailInfo();
    
    // Lade Send-Historie
    loadSendHistory();
}

async function loadEmailInfo() {
    try {
        console.log("Lade E-Mail Informationen...");
        addDebugLog("Starte E-Mail Laden...");
        
        // Warten bis Office.js geladen ist
        if (!Office.context || !Office.context.mailbox) {
            console.error("Office.context.mailbox nicht verfügbar");
            addErrorLog("Office.context.mailbox nicht verfügbar");
            showStatus("Office.js nicht geladen", "error");
            return;
        }
        
        addDebugLog("Office.context.mailbox verfügbar");
        
        // E-Mail Item abrufen
        const item = Office.context.mailbox.item;
        
        if (!item) {
            console.error("Kein E-Mail Item gefunden");
            addErrorLog("Kein E-Mail Item gefunden");
            showStatus("Keine E-Mail ausgewählt", "error");
            return;
        }
        
        console.log("E-Mail Item gefunden:", item);
        addDebugLog("E-Mail Item gefunden: " + JSON.stringify(item, null, 2));
        
        // E-Mail Metadaten sammeln (nur das Wichtigste)
        emailData = {
            id: item.itemId || "unknown",
            subject: item.subject || "Kein Betreff",
            sender: (item.from && item.from.emailAddress) ? item.from.emailAddress : "Unbekannt",
            senderName: (item.from && item.from.displayName) ? item.from.displayName : "Unbekannt",
            receivedTime: item.dateTimeCreated || new Date(),
            body: await getEmailBody(item)
        };
        
        console.log("E-Mail Daten:", emailData);
        
        currentEmail = emailData;
        
        // E-Mail Info anzeigen
        displayEmailInfo(emailData);
        
        // Gespeicherte Daten für diese E-Mail laden
        loadSavedEmailData(emailData.id);
        
    } catch (error) {
        console.error("Fehler beim Laden der E-Mail:", error);
        showStatus("Fehler beim Laden der E-Mail: " + error.message, "error");
        
        // Fallback: Zeige Test-Daten
        emailData = {
            id: "test-" + Date.now(),
            subject: "Test E-Mail",
            sender: "test@example.com",
            senderName: "Test Sender",
            receivedTime: new Date(),
            toRecipients: "empfaenger@example.com",
            ccRecipients: "",
            body: "Dies ist eine Test-E-Mail für das CRM Add-in.",
            attachments: 0,
            importance: "Normal",
            isRead: true
        };
        
        displayEmailInfo(emailData);
    }
}

async function getEmailBody(item) {
    return new Promise((resolve, reject) => {
        try {
            if (item.body) {
                item.body.getAsync(Office.CoercionType.Text, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("E-Mail Body geladen:", result.value ? result.value.substring(0, 100) + "..." : "Leer");
                        resolve(result.value || "");
                    } else {
                        console.error("Fehler beim Laden des E-Mail Body:", result.error);
                        resolve("");
                    }
                });
            } else {
                console.log("Kein E-Mail Body verfügbar");
                resolve("");
            }
        } catch (error) {
            console.error("Fehler in getEmailBody:", error);
            resolve("");
        }
    });
}

function displayEmailInfo(data) {
    const emailInfoDiv = document.getElementById('emailInfo');
    
    emailInfoDiv.innerHTML = `
        <div><strong>Von:</strong> ${data.senderName} (${data.sender})</div>
        <div><strong>Betreff:</strong> ${data.subject}</div>
        <div><strong>Empfangen:</strong> ${new Date(data.receivedTime).toLocaleString('de-DE')}</div>
    `;
}

function getImportanceText(importance) {
    // Fallback für Office.MailboxEnums falls nicht verfügbar
    if (typeof Office !== 'undefined' && Office.MailboxEnums && Office.MailboxEnums.ItemImportance) {
        switch(importance) {
            case Office.MailboxEnums.ItemImportance.High: return "Hoch";
            case Office.MailboxEnums.ItemImportance.Low: return "Niedrig";
            default: return "Normal";
        }
    } else {
        // Fallback ohne Office.MailboxEnums
        switch(importance) {
            case "High": return "Hoch";
            case "Low": return "Niedrig";
            case "Normal": return "Normal";
            default: return "Normal";
        }
    }
}

function selectRating(rating) {
    console.log("Rating ausgewählt:", rating);
    
    // Alle Rating Buttons deselektieren
    document.querySelectorAll('.rating-btn').forEach(btn => {
        btn.classList.remove('selected');
    });
    
    // Gewählten Button selektieren
    const selectedBtn = document.querySelector(`[data-rating="${rating}"]`);
    if (selectedBtn) {
        selectedBtn.classList.add('selected');
        console.log("Button selektiert:", selectedBtn);
    }
    
    // Rating in emailData speichern
    if (emailData) {
        emailData.rating = rating;
        console.log("Rating gespeichert:", emailData.rating);
    }
}

function loadSavedEmailData(emailId) {
    const savedData = localStorage.getItem(`email_${emailId}`);
    if (savedData) {
        const data = JSON.parse(savedData);
        document.getElementById('comment').value = data.comment || '';
        
        if (data.rating) {
            selectRating(data.rating);
        }
    }
}

async function sendToCRM() {
    if (!emailData) {
        showStatus("Keine E-Mail-Daten verfügbar", "error");
        return;
    }
    
    if (!emailData.rating) {
        showStatus("Bitte eine Bewertung auswählen", "error");
        return;
    }
    
    try {
        // Fester Webhook
        const webhookUrl = 'https://services.leadconnectorhq.com/hooks/mQuST3AEkqT3w3s1mfor/webhook-trigger/d7889f1c-5fbb-46fe-b720-2bcd9fab7c63';
        
        // Webhook-Daten für LeadConnector vorbereiten (nur das Wichtigste)
        const webhookData = {
            email: {
                id: emailData.id,
                subject: emailData.subject,
                sender_email: emailData.sender,
                sender_name: emailData.senderName,
                received_time: emailData.receivedTime,
                rating: emailData.rating,
                comment: document.getElementById('comment').value,
                processed_at: new Date().toISOString(),
                source: "Outlook Add-in CRM Manager"
            }
        };
        
        // Webhook senden
        const response = await fetch(webhookUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(webhookData)
        });
        
        if (response.ok) {
            showStatus("E-Mail erfolgreich an CRM gesendet!", "success");
            
            // E-Mail als verarbeitet markieren mit detailliertem Log
            const logEntry = {
                emailId: emailData.id,
                subject: emailData.subject,
                sender: emailData.sender,
                rating: emailData.rating,
                comment: document.getElementById('comment').value,
                processedAt: new Date().toISOString(),
                sentToCRM: true
            };
            
            localStorage.setItem(`email_${emailData.id}_processed`, JSON.stringify(logEntry));
            
            // Log zur Send-Historie hinzufügen
            const sendHistory = JSON.parse(localStorage.getItem('sendHistory') || '[]');
            sendHistory.unshift(logEntry);
            
            // Nur die letzten 50 Einträge behalten
            if (sendHistory.length > 50) {
                sendHistory.splice(50);
            }
            
            localStorage.setItem('sendHistory', JSON.stringify(sendHistory));
            
            addDebugLog(`E-Mail gesendet: ${emailData.subject} (ID: ${emailData.id})`);
            
            // Textfeld zurücksetzen
            document.getElementById('comment').value = '';
            
            // Rating zurücksetzen
            document.querySelectorAll('.rating-btn').forEach(btn => {
                btn.classList.remove('selected');
            });
            
            // Send-Historie aktualisieren
            loadSendHistory();
            
        } else {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
    } catch (error) {
        console.error("Fehler beim Senden an CRM:", error);
        showStatus("Fehler beim Senden an CRM: " + error.message, "error");
    }
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.innerHTML = `<div class="status ${type}">${message}</div>`;
    
    // Status nach 5 Sekunden ausblenden
    setTimeout(() => {
        statusDiv.innerHTML = '';
    }, 5000);
}

// Utility-Funktionen
function formatDate(dateString) {
    return new Date(dateString).toLocaleString('de-DE');
}

function truncateText(text, maxLength = 100) {
    return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
}

function loadSendHistory() {
    const sendHistoryList = document.getElementById('sendHistoryList');
    const sendHistory = JSON.parse(localStorage.getItem('sendHistory') || '[]');
    
    if (sendHistory.length > 0) {
        const historyHtml = sendHistory.slice(0, 10).map(entry => {
            const date = new Date(entry.processedAt).toLocaleString('de-DE');
            return `
                <div class="history-item">
                    <div class="history-date">${date}</div>
                    <div class="history-comment">${entry.comment || 'Kein Kommentar'}</div>
                </div>
            `;
        }).join('');
        
        sendHistoryList.innerHTML = historyHtml;
    } else {
        sendHistoryList.innerHTML = '<div class="history-note">Hinweis: Die Send-Historie ist cookie-basiert und daher nur temporär</div>';
    }
}

// Debug Functions
function setupDebugPanel() {
    const toggleBtn = document.getElementById('toggleDebug');
    const debugContent = document.getElementById('debugContent');
    const debugPanel = document.getElementById('debugPanel');
    
    toggleBtn.addEventListener('click', function() {
        if (debugPanel.style.display === 'none') {
            debugPanel.style.display = 'block';
            debugContent.style.display = 'block';
            toggleBtn.textContent = 'Ausblenden';
            updateDebugInfo();
        } else {
            debugPanel.style.display = 'none';
            debugContent.style.display = 'none';
            toggleBtn.textContent = 'Einblenden';
        }
    });
    
    // Console override für Debug-Logs
    const originalLog = console.log;
    const originalError = console.error;
    
    console.log = function(...args) {
        originalLog.apply(console, args);
        addDebugLog(args.join(' '));
    };
    
    console.error = function(...args) {
        originalError.apply(console, args);
        addErrorLog(args.join(' '));
    };
}

function addDebugLog(message) {
    const timestamp = new Date().toLocaleTimeString();
    debugLogs.push(`[${timestamp}] ${message}`);
    
    // Nur die letzten 20 Logs behalten
    if (debugLogs.length > 20) {
        debugLogs.shift();
    }
    
    updateDebugInfo();
}

function addErrorLog(message) {
    const timestamp = new Date().toLocaleTimeString();
    errorLogs.push(`[${timestamp}] ${message}`);
    
    // Nur die letzten 10 Fehler behalten
    if (errorLogs.length > 10) {
        errorLogs.shift();
    }
    
    updateDebugInfo();
}

function updateDebugInfo() {
    // Office.js Status
    const officeStatus = document.getElementById('officeStatus');
    if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
        officeStatus.innerHTML = '<span style="color: green;">✓ Office.js geladen</span>';
    } else {
        officeStatus.innerHTML = '<span style="color: red;">✗ Office.js nicht verfügbar</span>';
    }
    
    // E-Mail Item Status
    const emailItemStatus = document.getElementById('emailItemStatus');
    if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        emailItemStatus.innerHTML = '<span style="color: green;">✓ E-Mail Item verfügbar</span>';
    } else {
        emailItemStatus.innerHTML = '<span style="color: red;">✗ Kein E-Mail Item</span>';
    }
    
    // E-Mail Daten
    const emailDataDebug = document.getElementById('emailDataDebug');
    if (emailData) {
        emailDataDebug.textContent = JSON.stringify(emailData, null, 2);
    } else {
        emailDataDebug.textContent = 'Keine E-Mail Daten verfügbar';
    }
    
    // Console Logs
    const consoleLogs = document.getElementById('consoleLogs');
    consoleLogs.innerHTML = debugLogs.join('<br>');
    
    // Error Logs
    const errorLogsDiv = document.getElementById('errorLogs');
    errorLogsDiv.innerHTML = errorLogs.length > 0 ? errorLogs.join('<br>') : 'Keine Fehler';
    
    // Send History
    const sendHistoryDiv = document.getElementById('sendHistory');
    const sendHistory = JSON.parse(localStorage.getItem('sendHistory') || '[]');
    if (sendHistory.length > 0) {
        const historyHtml = sendHistory.slice(0, 10).map(entry => 
            `<div style="margin-bottom: 5px; font-size: 10px;">
                <strong>${entry.subject}</strong><br>
                Von: ${entry.sender}<br>
                Rating: ${entry.rating} | Zeit: ${new Date(entry.processedAt).toLocaleString('de-DE')}
            </div>`
        ).join('');
        sendHistoryDiv.innerHTML = historyHtml;
    } else {
        sendHistoryDiv.innerHTML = 'Keine gesendeten E-Mails';
    }
    
    // LocalStorage Keys
    const localStorageKeysDiv = document.getElementById('localStorageKeys');
    const keys = Object.keys(localStorage).filter(key => key.startsWith('email_') || key === 'sendHistory');
    localStorageKeysDiv.innerHTML = keys.length > 0 ? 
        keys.map(key => `<div style="font-size: 10px;">${key}</div>`).join('') : 
        'Keine E-Mail-Daten gespeichert';
}
