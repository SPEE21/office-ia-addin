// Contrôleur principal pour le volet Mammouth IA
let officeHelpers;
let mammouthClient;

let currentHost = "Browser";
let apiKey = "";
let selectedModel = "gpt-4o";
let isContextActive = false;

// Initialiser Office.js
Office.onReady((info) => {
  officeHelpers = new OfficeHelpers();
  mammouthClient = new MammouthClient();
  
  currentHost = officeHelpers.getHost();
  
  // Initialisation de l'interface utilisateur
  initApp();
});

/**
 * Configure les écouteurs d'événements, charge les paramètres et met à jour l'UI selon l'hôte.
 */
function initApp() {
  // Mettre à jour les badges d'hôte et les textes d'informations
  const hostBadge = document.getElementById("host-badge");
  const infoHostName = document.getElementById("info-host-name");
  const infoOfficeStatus = document.getElementById("info-office-status");

  infoHostName.textContent = currentHost;
  infoOfficeStatus.textContent = officeHelpers.isAvailable() ? "Actif (Office.js)" : "Inactif (Simulé)";

  if (currentHost === "Word") {
    hostBadge.textContent = "Mode Word";
    hostBadge.classList.add("word-mode");
    document.getElementById("assistants-word-view").classList.remove("hidden");
  } else if (currentHost === "Excel") {
    hostBadge.textContent = "Mode Excel";
    hostBadge.classList.add("excel-mode");
    document.getElementById("assistants-excel-view").classList.remove("hidden");
  } else {
    hostBadge.textContent = "Mode Navigateur";
    // Par défaut en navigateur, on affiche la vue Word pour démo
    document.getElementById("assistants-word-view").classList.remove("hidden");
  }

  // Initialiser les icônes Lucide
  lucide.createIcons();

  // Charger les paramètres stockés
  loadSettings();

  // Enregistrer les écouteurs d'événements
  setupEventListeners();

  // Test de connexion automatique si clé existante
  if (apiKey) {
    validateAndLoadModels(true);
  }
}

// =========================================================================
// GESTION DES PARAMÈTRES & CONFIGURATION
// =========================================================================

function loadSettings() {
  apiKey = localStorage.getItem("mammouth_api_key") || "";
  selectedModel = localStorage.getItem("mammouth_selected_model") || "gpt-4o";
  const theme = localStorage.getItem("mammouth_theme") || "dark";

  // Appliquer la clé API dans l'input
  document.getElementById("api-key-input").value = apiKey;

  // Appliquer le thème
  applyTheme(theme);

  // Mettre à jour l'indicateur visuel de clé manquante
  updateKeyWarningBanner();
}

function applyTheme(theme) {
  const body = document.body;
  const themeBtns = document.querySelectorAll(".theme-choice-btn");
  const toggleBtnIcon = document.querySelector("#theme-toggle-btn i");

  if (theme === "light") {
    body.classList.remove("dark-theme");
    body.classList.add("light-theme");
    if (toggleBtnIcon) toggleBtnIcon.setAttribute("data-lucide", "moon");
  } else {
    body.classList.remove("light-theme");
    body.classList.add("dark-theme");
    if (toggleBtnIcon) toggleBtnIcon.setAttribute("data-lucide", "sun");
  }

  // Activer le bon bouton dans les réglages
  themeBtns.forEach(btn => {
    if (btn.getAttribute("data-theme") === theme) {
      btn.classList.add("active");
    } else {
      btn.classList.remove("active");
    }
  });

  lucide.createIcons();
  localStorage.setItem("mammouth_theme", theme);
}

function updateKeyWarningBanner() {
  const banner = document.getElementById("key-warning-banner");
  const saveBtn = document.getElementById("save-api-key-btn");
  const sendBtn = document.getElementById("send-chat-btn");

  if (!apiKey) {
    banner.classList.remove("hidden");
    sendBtn.disabled = true;
  } else {
    banner.classList.add("hidden");
    sendBtn.disabled = false;
  }
}

/**
 * Valide la clé API entrée, charge les modèles et met à jour l'UI.
 */
async function validateAndLoadModels(isSilent = false) {
  const statusIndicator = document.getElementById("api-status-indicator");
  const statusText = statusIndicator.querySelector(".status-text");
  const modelSelect = document.getElementById("model-select");
  const refreshBtn = document.getElementById("refresh-models-btn");

  if (!apiKey) {
    statusIndicator.className = "api-status-tag unconfigured";
    statusText.textContent = "Non configuré";
    return;
  }

  try {
    if (!isSilent) {
      statusText.textContent = "Validation...";
      statusIndicator.className = "api-status-tag unconfigured";
    }

    // Valider la clé
    await mammouthClient.validateApiKey(apiKey);

    // Si valide, mettre à jour le statut
    statusIndicator.className = "api-status-tag configured";
    statusText.textContent = "Connecté";
    updateKeyWarningBanner();

    // Charger les modèles
    modelSelect.disabled = false;
    refreshBtn.disabled = false;
    modelSelect.innerHTML = '<option value="">Chargement...</option>';

    const models = await mammouthClient.getModels();
    modelSelect.innerHTML = "";

    models.forEach(model => {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.name || model.id;
      if (model.id === selectedModel) option.selected = true;
      modelSelect.appendChild(option);
    });

    // Mettre à jour l'indicateur de modèle dans le chat
    document.getElementById("chat-model-indicator").textContent = selectedModel;

  } catch (error) {
    console.error(error);
    statusIndicator.className = "api-status-tag unconfigured";
    statusText.textContent = "Erreur de clé";
    modelSelect.innerHTML = '<option value="error">Erreur de connexion</option>';
    modelSelect.disabled = true;
    refreshBtn.disabled = true;
    
    if (!isSilent) {
      alert("Erreur de connexion à Mammouth IA. Veuillez vérifier votre clé API.");
    }
  }
}

// =========================================================================
// CONFIGURATION DES ECOUTEURS D'EVENEMENTS
// =========================================================================

function setupEventListeners() {
  // Navigation par Onglets
  const tabs = document.querySelectorAll(".nav-tab");
  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      tab.classList.add("active");

      const targetId = tab.getAttribute("data-tab");
      document.querySelectorAll(".tab-pane").forEach(pane => {
        pane.classList.remove("active");
      });
      document.getElementById(targetId).classList.add("active");
    });
  });

  // Actionneur de Thème Rapide (Header)
  document.getElementById("theme-toggle-btn").addEventListener("click", () => {
    const currentTheme = document.body.classList.contains("light-theme") ? "dark" : "light";
    applyTheme(currentTheme);
  });

  // Actionneur de Thème dans Réglages
  document.querySelectorAll(".theme-choice-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      applyTheme(btn.getAttribute("data-theme"));
    });
  });

  // Raccourci vers Réglages depuis la bannière d'alerte
  document.getElementById("go-to-settings-btn").addEventListener("click", () => {
    document.querySelector('[data-tab="tab-settings"]').click();
  });

  // Afficher / masquer la clé API
  const togglePasswordBtn = document.getElementById("toggle-password-btn");
  const apiKeyInput = document.getElementById("api-key-input");
  togglePasswordBtn.addEventListener("click", () => {
    const isPassword = apiKeyInput.type === "password";
    apiKeyInput.type = isPassword ? "text" : "password";
    togglePasswordBtn.querySelector("i").setAttribute("data-lucide", isPassword ? "eye-off" : "eye");
    lucide.createIcons();
  });

  // Enregistrer la clé API
  document.getElementById("save-api-key-btn").addEventListener("click", () => {
    apiKey = apiKeyInput.value.trim();
    localStorage.setItem("mammouth_api_key", apiKey);
    validateAndLoadModels(false);
  });

  // Rafraîchir les modèles
  document.getElementById("refresh-models-btn").addEventListener("click", () => {
    validateAndLoadModels(false);
  });

  // Modifier le modèle préféré
  const modelSelect = document.getElementById("model-select");
  modelSelect.addEventListener("change", () => {
    selectedModel = modelSelect.value;
    localStorage.setItem("mammouth_selected_model", selectedModel);
    document.getElementById("chat-model-indicator").textContent = selectedModel;
  });

  // Auto-grow du textarea d'entrée chat
  const chatInput = document.getElementById("chat-input");
  chatInput.addEventListener("input", () => {
    chatInput.style.height = "auto";
    chatInput.style.height = (chatInput.scrollHeight) + "px";
    
    // Activer/Désactiver bouton d'envoi
    document.getElementById("send-chat-btn").disabled = (!chatInput.value.trim() || !apiKey);
  });

  // Envoyer message Chat
  document.getElementById("send-chat-btn").addEventListener("click", sendChatMessage);
  chatInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      if (chatInput.value.trim() && apiKey) sendChatMessage();
    }
  });

  // Suggestion Chips (Chat)
  document.querySelectorAll(".suggestion-chip").forEach(chip => {
    chip.addEventListener("click", () => {
      chatInput.value = chip.getAttribute("data-prompt");
      chatInput.dispatchEvent(new Event("input"));
      chatInput.focus();
    });
  });

  // Inclure la sélection comme contexte
  const contextToggle = document.getElementById("include-selection-btn");
  contextToggle.addEventListener("click", async () => {
    isContextActive = !isContextActive;
    
    if (isContextActive) {
      contextToggle.classList.add("active");
      document.getElementById("chat-context-indicator").classList.remove("hidden");
    } else {
      contextToggle.classList.remove("active");
      document.getElementById("chat-context-indicator").classList.add("hidden");
    }
  });

  document.getElementById("clear-context-btn").addEventListener("click", () => {
    isContextActive = false;
    contextToggle.classList.remove("active");
    document.getElementById("chat-context-indicator").classList.add("hidden");
  });

  // Fermer le panneau flottant de résultats d'assistants
  document.getElementById("btn-close-results").addEventListener("click", hideAssistantResults);

  // Actions de résultats d'assistants (Insérer, Copier, Remplacer)
  const resultTextElement = document.getElementById("assistant-results-text");
  
  document.getElementById("btn-results-copy").addEventListener("click", () => {
    navigator.clipboard.writeText(resultTextElement.textContent);
    alert("Copié dans le presse-papier !");
  });

  document.getElementById("btn-results-apply").addEventListener("click", async () => {
    const text = resultTextElement.textContent;
    if (currentHost === "Word") {
      await officeHelpers.insertTextWord(text, "replace");
    } else if (currentHost === "Excel") {
      await officeHelpers.writeExcelSelection(text, false);
    }
    hideAssistantResults();
  });

  document.getElementById("btn-results-insert").addEventListener("click", async () => {
    const text = resultTextElement.textContent;
    if (currentHost === "Word") {
      await officeHelpers.insertTextWord(text, "after");
    } else if (currentHost === "Excel") {
      // Pour Excel, insérer après signifie écrire sous la cellule
      alert("Veuillez sélectionner la cellule cible dans Excel pour y insérer la donnée.");
    }
    hideAssistantResults();
  });

  // =========================================================================
  // ACTIONS ASSISTANTS RAPIDES (WORD)
  // =========================================================================
  
  // Reformuler Word
  document.getElementById("btn-rewrite").addEventListener("click", async () => {
    const tone = document.getElementById("rewrite-tone").value;
    runWordAssistant("rewrite", `Reformule le texte suivant en utilisant un ton ${tone}. Reste fidèle au sens d'origine.`);
  });

  // Corriger Word
  document.getElementById("btn-correct").addEventListener("click", async () => {
    runWordAssistant("correct", "Corrige l'orthographe, la grammaire et la ponctuation du texte suivant. Renvoie uniquement le texte corrigé, sans introduction ni explications.");
  });

  // Résumer Word
  document.getElementById("btn-summarize").addEventListener("click", async () => {
    runWordAssistant("summarize", "Fais une synthèse structurée et claire sous forme de puces (bullets points) du texte suivant.");
  });

  // Continuer Rédaction Word
  document.getElementById("btn-continue").addEventListener("click", async () => {
    runWordAssistant("continue", "Rédige la suite logique du paragraphe suivant. Conserve le même style d'écriture.");
  });

  // Générer texte libre Word
  document.getElementById("btn-generate-text").addEventListener("click", async () => {
    const prompt = document.getElementById("prompt-generator").value.trim();
    if (!prompt) return alert("Veuillez décrire ce que vous souhaitez rédiger.");
    
    showAssistantResults();
    try {
      const messages = [{ role: "user", content: prompt }];
      let result = "";
      
      await mammouthClient.chatCompletion(apiKey, selectedModel, messages, (chunk) => {
        result += chunk;
        updateAssistantText(result);
      });
    } catch (err) {
      updateAssistantText(`Erreur lors de la génération: ${err.message}`);
    }
  });

  // =========================================================================
  // ACTIONS ASSISTANTS RAPIDES (EXCEL)
  // =========================================================================
  
  // Générer formule Excel
  document.getElementById("btn-generate-formula").addEventListener("click", async () => {
    const instruction = document.getElementById("formula-instruction").value.trim();
    if (!instruction) return alert("Veuillez décrire le calcul à effectuer.");
    
    const resultBox = document.getElementById("formula-result-container");
    const codeElement = document.getElementById("formula-code");
    
    resultBox.classList.remove("hidden");
    codeElement.textContent = "Génération de la formule...";
    
    try {
      const messages = [
        { 
          role: "system", 
          content: "Tu es un expert Excel. L'utilisateur va te décrire une opération à effectuer. Tu dois renvoyer UNIQUEMENT la formule Excel correspondante (ex: =SOMME(A1:A10)), de préférence en version internationale et française. Écris la formule brute sans bloc de code markdown." 
        },
        { role: "user", content: instruction }
      ];
      
      let formula = "";
      await mammouthClient.chatCompletion(apiKey, selectedModel, messages, (chunk) => {
        formula += chunk;
        codeElement.textContent = formula.trim();
      });

      // Nettoyer la formule (retirer les accents de blocs markdown éventuels)
      const cleanForm = cleanFormula(codeElement.textContent);
      codeElement.textContent = cleanForm;

      // Boutons de formules
      document.getElementById("btn-insert-formula").onclick = async () => {
        await officeHelpers.writeExcelSelection(cleanForm, true);
        resultBox.classList.add("hidden");
      };

      document.getElementById("btn-copy-formula").onclick = () => {
        navigator.clipboard.writeText(cleanForm);
        alert("Formule copiée !");
      };
      
    } catch (err) {
      codeElement.textContent = `Erreur: ${err.message}`;
    }
  });

  // Analyse de données Excel
  document.getElementById("btn-analyze-excel").addEventListener("click", async () => {
    showAssistantResults();
    try {
      updateAssistantText("Récupération de vos données...");
      const data = await officeHelpers.getSelectedExcelData();
      
      if (!data || !data.hasData) {
        updateAssistantText("Aucune donnée détectée dans votre sélection Excel. Veuillez sélectionner une plage de cellules contenant des valeurs.");
        return;
      }

      updateAssistantText(`Analyse en cours sur la plage ${data.address}...\n\n`);

      const messages = [
        { 
          role: "system", 
          content: "Tu es un analyste de données expert. L'utilisateur va te soumettre des données issues d'un tableau Excel. Fais une analyse approfondie : dégage les tendances clés, identifie les anomalies, propose des recommandations d'amélioration et suggère des types de visualisations (graphiques) pertinents." 
        },
        { role: "user", content: `Voici mes données sélectionnées :\n\n${data.markdown}` }
      ];

      let analysis = "";
      await mammouthClient.chatCompletion(apiKey, selectedModel, messages, (chunk) => {
        analysis += chunk;
        updateAssistantText(analysis);
      });
      
    } catch (err) {
      updateAssistantText(`Erreur lors de l'analyse : ${err.message}`);
    }
  });

  // Remplissage intelligent Excel
  document.getElementById("btn-fill-excel").addEventListener("click", async () => {
    const prompt = document.getElementById("fill-instruction").value.trim();
    if (!prompt) return alert("Veuillez décrire le schéma ou la consigne de remplissage.");

    showAssistantResults();
    try {
      updateAssistantText("Chargement des cellules...");
      const data = await officeHelpers.getSelectedExcelData();
      
      if (!data || !data.hasData) {
        updateAssistantText("Veuillez sélectionner la plage de données à remplir (incluant des lignes modèles complètes et des cases vides à remplir).");
        return;
      }

      updateAssistantText("L'IA calcule le schéma de remplissage...\n\n");

      const messages = [
        {
          role: "system",
          content: "Tu es une IA de complétion de tableurs. Tu dois analyser le tableau fourni et compléter les cases vides selon la consigne de l'utilisateur. Tu dois renvoyer UNIQUEMENT un tableau au format JSON brut représentant la grille complétée (un tableau 2D de chaînes/nombres) pour pouvoir l'insérer directement. Remplis uniquement les cellules vides. Ne renvoie rien d'autre que le tableau JSON valide."
        },
        {
          role: "user",
          content: `Consigne : ${prompt}\n\nDonnées actuelles (2D array):\n${JSON.stringify(data.values)}`
        }
      ];

      let jsonResponse = "";
      await mammouthClient.chatCompletion(apiKey, selectedModel, messages, (chunk) => {
        jsonResponse += chunk;
        updateAssistantText("Calcul du tableau JSON...\n\n" + jsonResponse);
      });

      // Tenter de parser le tableau
      try {
        const cleanJson = jsonResponse.replace(/```[a-zA-Z]*\n?/g, "").replace(/```/g, "").trim();
        const parsedGrid = JSON.parse(cleanJson);
        
        if (Array.isArray(parsedGrid) && parsedGrid.length === data.values.length) {
          updateAssistantText("Calcul terminé avec succès !");
          
          // Remplacer directement la sélection
          await officeHelpers.writeExcelSelection(parsedGrid, false);
          setTimeout(hideAssistantResults, 1000);
        } else {
          throw new Error("Format JSON invalide ou dimensions incorrectes.");
        }
      } catch (jsonErr) {
        updateAssistantText(`Erreur d'analyse JSON de la réponse. La réponse de l'IA n'était pas un tableau JSON exploitable pour Excel :\n\n${jsonResponse}`);
      }

    } catch (err) {
      updateAssistantText(`Erreur : ${err.message}`);
    }
  });
}

// =========================================================================
// LOGIQUE DE CHAT INTERACTIF
// =========================================================================

async function sendChatMessage() {
  const chatInput = document.getElementById("chat-input");
  const promptText = chatInput.value.trim();
  if (!promptText || !apiKey) return;

  // Vider le champ de saisie
  chatInput.value = "";
  chatInput.style.height = "auto";
  document.getElementById("send-chat-btn").disabled = true;

  // Créer et afficher le message utilisateur
  appendMessage(promptText, "user-message");

  // Préparer les messages pour l'API
  let messages = [];
  
  // 1. Ajouter le contexte de sélection si actif
  if (isContextActive) {
    try {
      let selectionText = "";
      if (currentHost === "Word") {
        selectionText = await officeHelpers.getSelectedTextWord();
        if (selectionText) {
          messages.push({
            role: "system",
            content: `Contexte du document Word sélectionné par l'utilisateur :\n"""\n${selectionText}\n"""`
          });
        }
      } else if (currentHost === "Excel") {
        const data = await officeHelpers.getSelectedExcelData();
        if (data && data.hasData) {
          messages.push({
            role: "system",
            content: `Contexte des données Excel sélectionnées par l'utilisateur :\n${data.markdown}`
          });
        }
      }
    } catch (err) {
      console.warn("Impossible de récupérer le contexte de sélection :", err);
    }
  }

  // 2. Ajouter l'historique récent (facultatif pour rester léger, ou on récupère les derniers messages du DOM)
  const chatMessagesDom = document.getElementById("chat-messages");
  const messageNodes = chatMessagesDom.querySelectorAll(".message:not(.system-message)");
  
  // Charger les 5 derniers messages pour le contexte conversationnel
  const startIdx = Math.max(0, messageNodes.length - 6);
  for (let i = startIdx; i < messageNodes.length; i++) {
    const node = messageNodes[i];
    const isUser = node.classList.contains("user-message");
    // Extraire le texte brut (en ignorant les boutons d'actions)
    const contentText = node.querySelector(".message-content").textContent;
    messages.push({
      role: isUser ? "user" : "assistant",
      content: contentText
    });
  }

  // 3. Ajouter le message utilisateur final s'il n'est pas déjà capturé
  messages.push({ role: "user", content: promptText });

  // Créer le message AI vide pour le streaming
  const aiMessageDiv = appendMessage("", "ai-message");
  const aiContentDiv = aiMessageDiv.querySelector(".message-content");
  
  // Ajouter l'animation de shimmer dans le message pendant le chargement
  aiContentDiv.innerHTML = `
    <div class="shimmer-loader">
      <div class="shimmer-line"></div>
      <div class="shimmer-line"></div>
    </div>
  `;

  try {
    let fullResponse = "";
    await mammouthClient.chatCompletion(apiKey, selectedModel, messages, (chunk) => {
      // Retirer le shimmer au premier chunk
      if (fullResponse === "") {
        aiContentDiv.innerHTML = "";
      }
      fullResponse += chunk;
      aiContentDiv.innerHTML = formatMarkdown(fullResponse);
      
      // Auto-scroll en bas de la boîte de chat
      chatMessagesDom.scrollTop = chatMessagesDom.scrollHeight;
    });

    // Ajouter les boutons d'actions pour le message IA
    addMessageActionButtons(aiMessageDiv, fullResponse);

  } catch (error) {
    aiContentDiv.innerHTML = `<span style="color: var(--color-red)">Erreur de communication : ${error.message}</span>`;
  }
}

/**
 * Ajoute un bloc de message dans le fil de discussion.
 */
function appendMessage(text, className) {
  const chatMessages = document.getElementById("chat-messages");
  
  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${className}`;
  
  const contentDiv = document.createElement("div");
  contentDiv.className = "message-content";
  contentDiv.innerHTML = className === "user-message" ? formatTextForHtml(text) : formatMarkdown(text);
  
  msgDiv.appendChild(contentDiv);
  chatMessages.appendChild(msgDiv);
  
  // Scroll
  chatMessages.scrollTop = chatMessages.scrollHeight;
  return msgDiv;
}

/**
 * Crée les boutons d'insertion et copie sous les messages de l'IA.
 */
function addMessageActionButtons(messageDiv, text) {
  const actionsDiv = document.createElement("div");
  actionsDiv.className = "message-actions";
  
  // Bouton Copier
  const copyBtn = document.createElement("button");
  copyBtn.className = "msg-action-btn";
  copyBtn.innerHTML = '<i data-lucide="copy"></i> Copier';
  copyBtn.onclick = () => {
    navigator.clipboard.writeText(text);
    copyBtn.innerHTML = '<i data-lucide="check"></i> Copié !';
    setTimeout(() => {
      copyBtn.innerHTML = '<i data-lucide="copy"></i> Copier';
      lucide.createIcons();
    }, 2000);
    lucide.createIcons();
  };

  // Bouton Insérer
  const insertBtn = document.createElement("button");
  insertBtn.className = "msg-action-btn";
  insertBtn.innerHTML = '<i data-lucide="plus"></i> Insérer';
  insertBtn.onclick = async () => {
    if (currentHost === "Word") {
      await officeHelpers.insertTextWord(text, "replace");
      insertBtn.innerHTML = '<i data-lucide="check"></i> Inséré !';
    } else if (currentHost === "Excel") {
      await officeHelpers.writeExcelSelection(text, false);
      insertBtn.innerHTML = '<i data-lucide="check"></i> Inséré !';
    } else {
      alert("Action disponible uniquement dans Word ou Excel.");
    }
    setTimeout(() => {
      insertBtn.innerHTML = '<i data-lucide="plus"></i> Insérer';
      lucide.createIcons();
    }, 2000);
    lucide.createIcons();
  };

  actionsDiv.appendChild(copyBtn);
  actionsDiv.appendChild(insertBtn);
  messageDiv.appendChild(actionsDiv);
  
  lucide.createIcons();
}

// =========================================================================
// LOGIQUE DES ASSISTANTS WORD
// =========================================================================

async function runWordAssistant(actionName, systemInstruction) {
  showAssistantResults();
  
  try {
    updateAssistantText("Lecture du texte sélectionné...");
    const selectedText = await officeHelpers.getSelectedTextWord();
    
    if (!selectedText || selectedText.trim() === "") {
      updateAssistantText("Aucune sélection détectée. Veuillez sélectionner du texte dans votre document Word et réessayer.");
      return;
    }

    updateAssistantText("Mammouth IA analyse et rédige...");
    
    const messages = [
      { role: "system", content: systemInstruction },
      { role: "user", content: selectedText }
    ];

    let fullOutput = "";
    await mammouthClient.chatCompletion(apiKey, selectedModel, messages, (chunk) => {
      fullOutput += chunk;
      updateAssistantText(fullOutput);
    });

  } catch (err) {
    updateAssistantText(`Erreur lors de l'application de l'assistant : ${err.message}`);
  }
}

// =========================================================================
// INTERFACE DE RESULTATS D'ASSISTANTS (FLOTANTE)
// =========================================================================

function showAssistantResults() {
  const panel = document.getElementById("assistant-results-wrapper");
  const textDiv = document.getElementById("assistant-results-text");
  const shimmer = panel.querySelector(".shimmer-loader");
  const actions = document.getElementById("assistant-results-actions");

  panel.classList.remove("hidden");
  textDiv.innerHTML = "";
  shimmer.classList.remove("hidden");
  actions.classList.add("hidden");
}

function hideAssistantResults() {
  document.getElementById("assistant-results-wrapper").classList.add("hidden");
}

function updateAssistantText(text) {
  const panel = document.getElementById("assistant-results-wrapper");
  const textDiv = document.getElementById("assistant-results-text");
  const shimmer = panel.querySelector(".shimmer-loader");
  const actions = document.getElementById("assistant-results-actions");

  shimmer.classList.add("hidden");
  textDiv.innerHTML = formatMarkdown(text);
  actions.classList.remove("hidden");
}

// =========================================================================
// UTILS & FORMATAGE DE TEXTES
// =========================================================================

function formatTextForHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
}

/**
 * Convertit un texte markdown minimal en code HTML propre pour l'affichage de l'IA.
 */
function formatMarkdown(text) {
  let html = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  // Gras
  html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");

  // Code Block
  html = html.replace(/```[a-zA-Z]*\n([\s\S]*?)\n```/g, "<pre><code>$1</code></pre>");

  // Inline Code
  html = html.replace(/`(.*?)`/g, "<code>$1</code>");

  // Retours à la ligne
  html = html.replace(/\n/g, "<br>");

  return html;
}

/**
 * Nettoie une formule Excel renvoyée pour n'avoir que du texte brut exécutable.
 */
function cleanFormula(formulaText) {
  let clean = formulaText.trim();
  
  // Retirer les blocs de code markdown (ex: ```excel ... ```)
  clean = clean.replace(/```[a-zA-Z]*\n?/g, "");
  clean = clean.replace(/```/g, "");
  clean = clean.trim();

  // Forcer le signe égal au début
  if (!clean.startsWith("=")) {
    clean = "=" + clean;
  }
  
  return clean;
}
