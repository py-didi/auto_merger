const App = {
    showLoader(message) {
        if(message) document.getElementById("loading-text").innerText = message;
        document.getElementById("loading-overlay").classList.remove("hidden");
        document.getElementById("loading-overlay").classList.add("flex");

        // РО ВЕР ВКЛЮЧАЕТ РЕЖИМ ЧТЕНИЯ КНИГИ
        if (typeof ModeRover !== 'undefined') {
            ModeRover.setWorkingState(true);
        }
    },

    hideLoader() {
        document.getElementById("loading-overlay").classList.remove("flex");
        document.getElementById("loading-overlay").classList.add("hidden");

        // РОВЕР ВОЗВРАЩАЕТСЯ В ОЖИДАНИЕ
        if (typeof ModeRover !== 'undefined') {
            ModeRover.setWorkingState(false);
        }
    },

    toggleCheckboxes(containerId, state) {
        const checkboxes = document.querySelectorAll(`#${containerId} input[type="checkbox"]`);
        checkboxes.forEach(cb => cb.checked = state);
    },

    switchMode(modeNum) {
        const tab1 = document.getElementById("nav-mode1");
        const tab2 = document.getElementById("nav-mode2");
        const scr1 = document.getElementById("screen-mode1");
        const scr2 = document.getElementById("screen-mode2");

        // Сбрасываем стили на дефолтные Win95
        tab1.className = "win95-nav-btn";
        tab2.className = "win95-nav-btn";
        
        scr1.classList.add("hidden");
        scr2.classList.add("hidden");

        // Применяем активный класс
        if (modeNum === 1) {
            tab1.classList.add("active-nav");
            scr1.classList.remove("hidden");
        } else {
            tab2.classList.add("active-nav");
            scr2.classList.remove("hidden");
        }
    },

    toggleInstructions(show) {
        const modal = document.getElementById("instruction-modal");
        if (show) modal.classList.remove("hidden"), modal.classList.add("flex");
        else modal.classList.add("hidden"), modal.classList.remove("flex");
    }
};