let deferredPrompt = null;
let hasReloadedForNewWorker = false;

function wireInstallButton(button) {
    if (!button) return;

    button.hidden = false;
    button.addEventListener("click", async () => {
        if (!deferredPrompt) return;
        deferredPrompt.prompt();
        await deferredPrompt.userChoice;
        deferredPrompt = null;
        button.hidden = true;
    });
}

window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    deferredPrompt = event;
    wireInstallButton(document.getElementById("install-trigger"));
    wireInstallButton(document.getElementById("install-cta"));
});

window.addEventListener("appinstalled", () => {
    deferredPrompt = null;
    const installButton = document.getElementById("install-trigger");
    const installCta = document.getElementById("install-cta");
    if (installButton) installButton.hidden = true;
    if (installCta) installCta.hidden = true;
});

if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
        navigator.serviceWorker.register("/service-worker.js?v=beta-5").then((registration) => {
            registration.update().catch(() => {});
        }).catch(() => {});
    });

    navigator.serviceWorker.addEventListener("controllerchange", () => {
        if (hasReloadedForNewWorker) return;
        hasReloadedForNewWorker = true;
        window.location.reload();
    });
}

document.addEventListener("submit", (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement)) return;
    const button = form.querySelector('button[type="submit"]');
    if (!button || button.dataset.submitting === "true") return;
    button.dataset.submitting = "true";
    button.dataset.originalText = button.textContent || "";
    button.textContent = "Aguarde...";
    button.disabled = true;
});
