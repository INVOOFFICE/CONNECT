async function apiCall(action, payload = {}) {
    showLoading(true);
    try {
        const params = new URLSearchParams({ pwa_secret: CONFIG.PWA_SECRET, action: action, ...payload });
        const response = await fetch(`${CONFIG.API_URL}?${params.toString()}`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        if (!response.ok) throw new Error('Erreur reseau');
        const data = await response.json();
        if (!data.ok) throw new Error(data.error || 'Reponse API invalide');
        return data;
    } catch (error) {
        showToast('Erreur: ' + error.message, 'error');
        throw error;
    } finally {
        showLoading(false);
    }
}

async function apiPost(action, data = {}) {
    showLoading(true);
    try {
        const payload = btoa(unescape(encodeURIComponent(JSON.stringify({
            pwa_secret: CONFIG.PWA_SECRET,
            action: action,
            ...data
        }))));
        const response = await fetch(`${CONFIG.API_URL}?pwa_secret=${encodeURIComponent(CONFIG.PWA_SECRET)}&payload=${encodeURIComponent(payload)}`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        if (!response.ok) throw new Error('Erreur reseau');
        const result = await response.json();
        if (!result.ok) throw new Error(result.error || 'Reponse API invalide');
        return result;
    } catch (error) {
        showToast('Erreur: ' + error.message, 'error');
        throw error;
    } finally {
        showLoading(false);
    }
}
