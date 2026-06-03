/**
 * Client d'intégration pour l'API Mammouth IA.
 * Fournit des méthodes pour communiquer avec les services d'IA compatibles OpenAI de Mammouth.
 */
class MammouthClient {
  constructor() {
    this.baseUrl = "https://api.mammouth.ai/v1";
    this.publicModelsUrl = "https://api.mammouth.ai/public/models";
  }

  /**
   * Valide la clé API en effectuant une requête de test légère.
   * @param {string} apiKey - La clé API à tester.
   * @returns {Promise<boolean>} True si la clé est valide, sinon lance une erreur.
   */
  async validateApiKey(apiKey) {
    if (!apiKey || apiKey.trim() === "") {
      throw new Error("Clé API manquante ou vide.");
    }

    try {
      const response = await fetch(`${this.baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: "gpt-4o", // Modèle standard présent chez Mammouth
          messages: [{ role: "user", content: "ping" }],
          max_tokens: 1
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        const errorMessage = errorData.error?.message || `Erreur HTTP ${response.status}`;
        throw new Error(errorMessage);
      }

      return true;
    } catch (error) {
      console.error("[MammouthClient] Erreur de validation de clé API:", error);
      throw error;
    }
  }

  /**
   * Récupère la liste des modèles disponibles.
   * @returns {Promise<Array>} Liste d'objets modèles.
   */
  async getModels() {
    try {
      // 1. Essai de l'API publique
      const response = await fetch(this.publicModelsUrl);
      if (response.ok) {
        const data = await response.json();
        // Si l'API renvoie directement un tableau ou une structure spécifique
        if (Array.isArray(data)) return data;
        if (data.data && Array.isArray(data.data)) return data.data;
        if (data.models && Array.isArray(data.models)) return data.models;
      }
    } catch (e) {
      console.warn("[MammouthClient] Impossible de charger via l'endpoint public, essai via v1/models...", e);
    }

    // Liste de repli (fallbacks) avec les modèles populaires si les endpoints échouent ou en cas de déconnexion
    return [
      { id: "gpt-4o", name: "GPT-4o (OpenAI)" },
      { id: "claude-3-5-sonnet", name: "Claude 3.5 Sonnet (Anthropic)" },
      { id: "gemini-1.5-pro", name: "Gemini 1.5 Pro (Google)" },
      { id: "mistral-large", name: "Mistral Large (Mistral)" },
      { id: "llama-3-1-405b", name: "Llama 3.1 405B (Meta)" },
      { id: "deepseek-coder", name: "DeepSeek Coder (DeepSeek)" }
    ];
  }

  /**
   * Exécute une complétion de chat avec streaming de la réponse.
   * @param {string} apiKey - Clé API de l'utilisateur.
   * @param {string} model - Identifiant du modèle IA à utiliser.
   * @param {Array} messages - Tableau de messages [{role: 'user', content: '...'}]
   * @param {function} onChunk - Callback appelé pour chaque nouveau fragment de texte reçu.
   * @returns {Promise<string>} La réponse complète construite.
   */
  async chatCompletion(apiKey, model, messages, onChunk) {
    if (!apiKey) throw new Error("Clé API non configurée.");

    try {
      const response = await fetch(`${this.baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: model,
          messages: messages,
          stream: true
        })
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(errText || `Erreur serveur: ${response.status}`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");
      let buffer = "";
      let fullText = "";

      while (true) {
        const { value, done } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop(); // Conserver la ligne incomplète restante

        for (const line of lines) {
          const cleanLine = line.trim();
          if (!cleanLine) continue;
          if (cleanLine === "data: [DONE]") continue;

          if (cleanLine.startsWith("data: ")) {
            try {
              const json = JSON.parse(cleanLine.substring(6));
              const chunk = json.choices[0]?.delta?.content || "";
              if (chunk) {
                fullText += chunk;
                onChunk(chunk);
              }
            } catch (err) {
              // Ignorer les erreurs d'analyse JSON pour les lignes fragmentées
            }
          }
        }
      }

      // Vider le buffer restant si nécessaire
      if (buffer && buffer.startsWith("data: ")) {
        try {
          const json = JSON.parse(buffer.substring(6));
          const chunk = json.choices[0]?.delta?.content || "";
          if (chunk) {
            fullText += chunk;
            onChunk(chunk);
          }
        } catch (e) {}
      }

      return fullText;
    } catch (error) {
      console.error("[MammouthClient] Erreur de complétion de chat:", error);
      throw error;
    }
  }
}

// Exposer pour les scripts du navigateur / d'Office
window.MammouthClient = MammouthClient;
