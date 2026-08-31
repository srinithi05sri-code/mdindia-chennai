document.addEventListener("DOMContentLoaded", () => {

    const saveAllBtn = document.getElementById("saveAllBtn");

    if (!saveAllBtn) {
        return;
    }

    saveAllBtn.addEventListener("click", async () => {

        // Prevent double click
        if (saveAllBtn.disabled) {
            return;
        }

        saveAllBtn.disabled = true;
        saveAllBtn.textContent = "Saving...";

        const rows = document.querySelectorAll(
            ".claims-table tbody tr"
        );

        const claims = [];

        rows.forEach(row => {

            const claimType =
                row.querySelector(".claim-type");

            const claimStatus =
                row.querySelector(".claim-status");

            const userRemark =
                row.querySelector(".user-remark");

            if (!claimType || !claimStatus || !userRemark) {
                return;
            }

            claims.push({
                id: claimType.dataset.id,
                claim_type: claimType.value,
                claim_status: claimStatus.value,
                user_remark: userRemark.value
            });

        });

        try {

            const response = await fetch(
                "/user/save-claims",
                {
                    method: "POST",

                    headers: {
                        "Content-Type": "application/json"
                    },

                    body: JSON.stringify({
                        claims: claims
                    })
                }
            );

            const result = await response.json();

            if (!response.ok || !result.success) {
                throw new Error(
                    result.message || "Save failed"
                );
            }

            alert("All claims saved successfully!");

        } catch (error) {

            console.error("SAVE ERROR:", error);

            alert(
                "Save failed: " + error.message
            );

        } finally {

            saveAllBtn.disabled = false;
            saveAllBtn.textContent = "Save All";
        }

    });

});
async function saveClaims() {

    const rows = document.querySelectorAll("#claimsTable tbody tr");

    const claims = [];

    rows.forEach(row => {

        claims.push({
            id: row.dataset.claimId,
            claim_type: row.querySelector(".claim-type")?.value || "",
            claim_status: row.querySelector(".claim-status")?.value || "Pending",
            user_remark: row.querySelector(".user-remark")?.value || ""
        });

    });

    try {

        const response = await fetch("/user/save", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                claims: claims
            })
        });

        const data = await response.json();

        if (!response.ok || !data.success) {
            throw new Error(data.message || "Save failed");
        }

        alert("Saved successfully!");

    } catch (error) {

        console.error(error);

        alert("Save failed: " + error.message);
    }
}