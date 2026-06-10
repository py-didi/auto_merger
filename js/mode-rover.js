const ModeRover = {
    phrases: [
        "🐶 Гав! Помните: Wildberries может поменять оферту три раза, пока вы пьете кофе. Проверяйте новости каждый час!",
        "🐶 Совет от Ровера: Если Ozon закинул карточку в скрытые, проверьте индекс цен. Возможно, робот нашел ваш товар дешевле на Авито.",
        "🐶 Гав! Гав! Никогда не клейте штрихкоды, которые распечатаны на умирающем картридже. Склад WB отправит всё в 'обезличку' и выставит счет.",
        "🐶 Ходят слухи, если настроить внутреннюю рекламу на WB ровно в 4 часа утра, ставка будет ниже на 15 рублей... Но это не точно. Гав!",
        "🐶 Лайфхак: Делаете flat-склейку? Убедитесь, что артикул поставщика совпадает до каждого символа, иначе базы сойдут с ума быстрее вас.",
        "🐶 Гав! Если Честный Знак опять завис — не паникуйте. Заварите чай. Против лома нет приема, против серверов ЦРПТ — тем более.",
        "🐶 Менеджер! А ты проверил габариты упаковки перед отправкой на склад? Коэффициент логистики х10 за неверные замеры ждет невнимательных!",
        "🐶 FBS-селлерам на заметку: не ждите дедлайна отгрузки. Статус 'Опоздание' прилетает на Ozon мгновенно и режет вам рейтинг.",
        "🐶 Гав! Потерялись остатки на складе? Проверьте вкладку 'Акты расхождений', маркетплейсы любят находить излишки через месяц."
    ],

    init() {
        this.sayRandom();
        // Ровер подкидывает базу каждые 20 секунд
        setInterval(() => this.sayRandom(), 20000);
    },

    sayRandom() {
        const cloud = document.getElementById("rover-text-cloud");
        if (!cloud) return;
        
        let currentText = cloud.innerText;
        let available = this.phrases.filter(p => p !== currentText);
        let randomPhrase = available[Math.floor(Math.random() * available.length)];
        
        cloud.innerText = randomPhrase;
    },

    // Переключение гифок в зависимости от загрузки процессора
    setWorkingState(isWorking) {
        const idleGif = document.getElementById("rover-gif-idle");
        const workingGif = document.getElementById("rover-gif-working");
        const cloud = document.getElementById("rover-text-cloud");
        
        if (!idleGif || !workingGif) return;

        if (isWorking) {
            idleGif.classList.add("hidden");
            workingGif.classList.remove("hidden");
            if (cloud) cloud.innerText = "🐶 Гав! Так, смотрю в книгу — вижу... А, нет, всё нормально, сейчас всё зальем и склеим!";
        } else {
            workingGif.classList.add("hidden");
            idleGif.classList.remove("hidden");
            this.sayRandom();
        }
    }
};

// Автозапуск пёселя при загрузке скрипта
if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => ModeRover.init());
} else {
    ModeRover.init();
}