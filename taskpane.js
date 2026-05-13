function saveData() {

    const data = {
        project: document.getElementById("project").value,
        client: document.getElementById("client").value,
        hours: document.getElementById("hours").value
    };

    Office.context.mailbox.item.loadCustomPropertiesAsync(function(result) {

        const props = result.value;

        props.set("reportingData", JSON.stringify(data));

        props.saveAsync(function() {
            alert("✅ Записано");
        });

    });
}


// Зареждане при отваряне
Office.onReady(function() {

    Office.context.mailbox.item.loadCustomPropertiesAsync(function(result) {

        const props = result.value;

        let data = {};

        try {
            data = JSON.parse(props.get("reportingData") || "{}");
        } catch(e) {}

        document.getElementById("project").value = data.project || "";
        document.getElementById("client").value = data.client || "";
        document.getElementById("hours").value = data.hours || "";

    });
});
``
