const PARTIALS = [
    'components/login.html',
    'components/splash.html',
    'sections/app-shell.html',
    'components/bottom-nav.html',
    'components/car-modal.html',
    'components/reservation-modal.html',
    'components/sync-modal.html',
    'components/settings-modal.html',
    'components/toast-container.html',
    'components/loading-overlay.html',
    'components/install-prompt.html'
];

async function loadPartials() {
    const root = document.getElementById('app-root');
    const fragments = await Promise.all(PARTIALS.map(async (partialPath) => {
        const response = await fetch(partialPath);
        if (!response.ok) throw new Error(`Impossible de charger ${partialPath}`);
        return response.text();
    }));
    root.innerHTML = fragments.join('\n');
}

document.addEventListener('DOMContentLoaded', async () => {
    try {
        await loadPartials();
        initializeApp();
    } catch (error) {
        document.body.innerHTML = `<div style="padding:24px;font-family:Arial,sans-serif;color:#0f172a;"><h1>Erreur de chargement</h1><p>${error.message}</p></div>`;
        console.error(error);
    }
});
