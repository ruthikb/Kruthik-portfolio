document.getElementById("contactForm").addEventListener("submit", function(event) {
      event.preventDefault();

      // Form Validation
      const name = document.getElementById("name").value.trim();
      const email = document.getElementById("email").value.trim();
      const subject = document.getElementById("subject").value.trim();
      const message = document.getElementById("message").value.trim();

      if (!name || !email || !subject || !message) {
        alert("⚠️ Please fill all the fields.");
        return;
      }

      const emailPattern = /^[^ ]+@[^ ]+\.[a-z]{2,3}$/;
      if (!email.match(emailPattern)) {
        alert("❌ Please enter a valid email address.");
        return;
      }

      // Prepare Data for Excel
      const data = [
        ["Name", "Email", "Subject", "Message"],
        [name, email, subject, message]
      ];

      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Contact_Data");

      const excelFile = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      saveAs(new Blob([excelFile], { type: "application/octet-stream" }), "ContactData.xlsx");

      alert("✅ Message saved to Excel successfully!");
      document.getElementById("contactForm").reset();
    });