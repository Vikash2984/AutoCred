document.addEventListener('DOMContentLoaded', () => {
    const genType = document.getElementById("gen-type");
    const dataPanel = document.getElementById("data-ingestion");
    const form = document.getElementById("certificate-form");

    const singleInput = `
        <div class="form-group">
            <label for="student-name">Student Name</label>
            <input type="text" id="student-name" name="student_name" required>
        </div>
        <div class="form-group">
            <label for="email">Email</label>
            <input type="email" id="email" name="email" required>
        </div>
        <div class="form-group">
            <label for="department">Department</label>
            <input type="text" id="department" name="department" required>
        </div>
        <div class="form-group">
            <label for="year">Academic Year</label>
            <select id="year" name="year" required>
                <option value="First">First</option>
                <option value="Second">Second</option>
                <option value="Third">Third</option>
                <option value="Fourth">Fourth</option>
            </select>
        </div>`;

    const bulkInput = `
        <div class="file-upload">
            <label for="bulk-upload" class="file-label">
                <span>Upload your .xlsx file</span>
                <input type="file" id="bulk-upload" name="file" accept=".xlsx" required>
            </label>
            <p class="file-name"></p>
        </div>`;

    // Initialize with bulk input
    dataPanel.innerHTML = bulkInput;

    genType.addEventListener("change", () => {
        if (genType.value === "bulk") {
            dataPanel.innerHTML = bulkInput;
        } else if (genType.value === "single") {
            dataPanel.innerHTML = singleInput;
        }
        addFileUploadListener();
    });

    function addFileUploadListener() {
        const fileUpload = document.getElementById("bulk-upload");
        const fileName = document.querySelector(".file-name");
        if (fileUpload) {
            fileUpload.addEventListener("change", (e) => {
                if (e.target.files.length > 0) {
                    fileName.textContent = `File selected: ${e.target.files[0].name}`;
                } else {
                    fileName.textContent = "";
                }
            });
        }
    }

    addFileUploadListener();

    function formatDate(inputDate) {
        const [year, month, day] = inputDate.split("-");
        return `${day}-${month}-${year}`;
    }

    form.addEventListener("submit", async (e) => {
        e.preventDefault();
        const formData = new FormData(form);
        const submitBtn = form.querySelector('.submit-btn');
        const originalBtnText = submitBtn.textContent;

        const eventDateInput = document.getElementById("event-date").value;
        if (eventDateInput) {
            const formattedDate = formatDate(eventDateInput);
            formData.set("event_date", formattedDate);
        }

        submitBtn.disabled = true;
        submitBtn.innerHTML = 'Generating... <span class="loading"></span>';

        const response = await fetch("http://127.0.0.1:8000/generate-certificates", {
            method: "POST",
            body: formData,
        });

        const result = await response.json();
        alert(result.message || "Certificates generated successfully!");

        submitBtn.disabled = false;
        submitBtn.textContent = originalBtnText;
    });
});
