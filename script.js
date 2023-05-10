document.getElementById("reportForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const startDate = document.getElementById("start_date").value;
    const endDate = document.getElementById("end_date").value;

    const response = await fetch("http://localhost:5000/generate_report", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify({
            start_date: startDate,
            end_date: endDate,
        }),
    });

    if (response.ok) {
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "output.xlsx";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    } else {
        alert("Error generating report.");
    }
});
