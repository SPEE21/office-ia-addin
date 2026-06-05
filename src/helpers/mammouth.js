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
    let apiModels = [];
    try {
      // 1. Essai de l'API publique
      const response = await fetch(this.publicModelsUrl);
      if (response.ok) {
        const data = await response.json();
        // Si l'API renvoie directement un tableau ou une structure spécifique
        if (Array.isArray(data)) apiModels = data;
        else if (data.data && Array.isArray(data.data)) apiModels = data.data;
        else if (data.models && Array.isArray(data.models)) apiModels = data.models;
      }
    } catch (e) {
      console.warn("[MammouthClient] Impossible de charger via l'endpoint public, essai via v1/models...", e);
    }

    // Si on n'a rien récupéré, utiliser les modèles de repli (fallbacks)
    if (apiModels.length === 0) {
      apiModels = [
        { id: "gpt-4o-mini", name: "GPT-4o Mini (OpenAI)" },
        { id: "gpt-4o", name: "GPT-4o (OpenAI)" },
        { id: "claude-3-5-haiku", name: "Claude 3.5 Haiku (Anthropic)" },
        { id: "claude-3-5-sonnet", name: "Claude 3.5 Sonnet (Anthropic)" },
        { id: "gemini-1.5-flash", name: "Gemini 1.5 Flash (Google)" },
        { id: "gemini-1.5-pro", name: "Gemini 1.5 Pro (Google)" },
        { id: "mistral-nemo", name: "Mistral Nemo (Mistral)" },
        { id: "mistral-large", name: "Mistral Large (Mistral)" },
        { id: "llama-3-1-8b", name: "Llama 3.1 8B (Meta)" },
        { id: "llama-3-1-405b", name: "Llama 3.1 405B (Meta)" },
        { id: "deepseek-coder", name: "DeepSeek Coder (DeepSeek)" },
        { id: "dall-e-3", name: "DALL-E 3 (OpenAI)" },
        { id: "dall-e-2", name: "DALL-E 2 (OpenAI)" },
        { id: "flux-schnell", name: "Flux Schnell (Black Forest Labs)" },
        { id: "stable-diffusion-xl", name: "Stable Diffusion XL (Stability AI)" }
      ];
    } else {
      // S'assurer de toujours inclure les modèles d'images s'ils ne sont pas listés par l'API
      const hasImageModel = apiModels.some(m => {
        const id = (m.id || "").toLowerCase();
        return id.includes("dall-e") || id.includes("flux") || id.includes("diffusion") || id.includes("sdxl") || id.includes("midjourney");
      });
      if (!hasImageModel) {
        apiModels.push(
          { id: "dall-e-3", name: "DALL-E 3 (OpenAI)" },
          { id: "dall-e-2", name: "DALL-E 2 (OpenAI)" },
          { id: "flux-schnell", name: "Flux Schnell (Black Forest Labs)" },
          { id: "stable-diffusion-xl", name: "Stable Diffusion XL (Stability AI)" }
        );
      }
    }

    return apiModels;
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

  /**
   * Génère une image à partir d'un prompt.
   * @param {string} apiKey - Clé API de l'utilisateur.
   * @param {string} model - Identifiant du modèle d'image (ex: dall-e-3).
   * @param {string} prompt - Description de l'image.
   * @returns {Promise<string>} Base64 de l'image générée (sans préfixe data URL).
   */
  async generateImage(apiKey, model, prompt) {
    if (!apiKey) throw new Error("Clé API non configurée.");

    try {
      const response = await fetch(`${this.baseUrl}/images/generations`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: model,
          prompt: prompt,
          n: 1,
          size: "1024x1024",
          response_format: "b64_json"
        })
      });

      if (!response.ok) {
        const errText = await response.text();
        console.warn("[MammouthClient] Échec de la génération en b64_json, essai en format standard (URL)...", errText);
        
        const fallbackResponse = await fetch(`${this.baseUrl}/images/generations`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${apiKey}`
          },
          body: JSON.stringify({
            model: model,
            prompt: prompt,
            n: 1,
            size: "1024x1024"
          })
        });
        
        if (!fallbackResponse.ok) {
          const fallbackErr = await fallbackResponse.text();
          throw new Error(fallbackErr || `Erreur serveur: ${fallbackResponse.status}`);
        }
        
        const data = await fallbackResponse.json();
        const imageUrl = data.data?.[0]?.url;
        if (!imageUrl) throw new Error("Aucune image renvoyée par le serveur.");
        
        return await this.convertImageUrlToBase64(imageUrl);
      }

      const data = await response.json();
      const b64Data = data.data?.[0]?.b64_json;
      if (b64Data) {
        return b64Data;
      }
      
      const imageUrl = data.data?.[0]?.url;
      if (imageUrl) {
        return await this.convertImageUrlToBase64(imageUrl);
      }
      
      throw new Error("Aucune donnée d'image renvoyée par le serveur.");
    } catch (error) {
      console.error("[MammouthClient] Erreur de génération d'image:", error);
      throw error;
    }
  }

  /**
   * Télécharge une image depuis une URL et la convertit en base64 brut.
   * @param {string} url - URL de l'image.
   * @returns {Promise<string>} Chaîne base64 brute.
   */
  async convertImageUrlToBase64(url) {
    const res = await fetch(url);
    const blob = await res.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64data = reader.result;
        const rawBase64 = base64data.split(",")[1];
        resolve(rawBase64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }
}

// Exposer pour les scripts du navigateur / d'Office
window.MammouthClient = MammouthClient;
