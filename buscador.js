(function() {
    const targetId = window.targetMatchId || 'match_1'; 

    const realizarSalto = () => {
        // Buscamos en el documento padre (donde está el texto)
        const principal = window.parent.document;
        const elemento = principal.getElementById(targetId);
        
        if (elemento) {
            elemento.scrollIntoView({
                behavior: 'smooth',
                block: 'center'
            });

            // Efecto visual para confirmar que es esa palabra
            elemento.style.outline = "4px solid #C0392B";
            elemento.style.transition = "outline 0.5s ease";
            setTimeout(() => { 
                elemento.style.outline = "none"; 
            }, 1000);
            return true;
        }
        return false;
    };

    // Intentar saltar
    if (!realizarSalto()) {
        const observer = new MutationObserver((mutations, obs) => {
            if (realizarSalto()) obs.disconnect();
        });
        observer.observe(window.parent.document.body, { childList: true, subtree: true });
        setTimeout(() => observer.disconnect(), 2000);
    }
})();






/* buscador.js
(function() {
    // Buscamos el ID que el input de Streamlit nos diga. 
    // Lo sacamos de un atributo que inyectaremos o de una variable global.
    const targetId = window.targetMatchId || 'match_1'; 

    const realizarSalto = () => {
        const principal = window.parent.document;
        const elemento = principal.getElementById(targetId);
        
        if (elemento) {
            elemento.scrollIntoView({
                behavior: 'smooth',
                block: 'center'
            });

            // Feedback visual potente
            elemento.style.outline = "4px solid #C0392B";
            elemento.style.backgroundColor = "#FFEB3B";
            elemento.style.transition = "all 0.5s ease";
            setTimeout(() => { 
                elemento.style.outline = "none"; 
            }, 1500);
            return true;
        }
        return false;
    };

    if (!realizarSalto()) {
        const observer = new MutationObserver((mutations, obs) => {
            if (realizarSalto()) obs.disconnect();
        });
        observer.observe(window.parent.document.body, { childList: true, subtree: true });
        setTimeout(() => observer.disconnect(), 3000);
    }
})(); */









