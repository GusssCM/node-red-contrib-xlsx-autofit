module.exports = function (RED) {
    const XLSX = require("xlsx");
    const fs = require("fs");
    const path = require("path");

    function XlsxAutofitNode(config) {
        RED.nodes.createNode(this, config);
        let node = this;

        node.on("input", function (msg) {
            try {
                let data = msg.payload;

                if (!Array.isArray(data) || data.length === 0) {
                    node.error("⚠️ Payload must be a non-empty array (2D array or array of objects)", msg);
                    return;
                }

                // 🔹 Convert array of objects to array of arrays (table format)
                if (data.length > 0 && typeof data[0] === "object" && !Array.isArray(data[0])) {
                    const headers = [...new Set(data.flatMap(obj => Object.keys(obj)))]; // Handle missing keys
                    data = [headers, ...data.map(obj => headers.map(h => obj[h] || ""))];
                }

                let ws = XLSX.utils.aoa_to_sheet(data);

                // 🔹 Autofit column width (Smarter width calculation)
                let colWidths = data[0].map((_, i) => {
                    let maxLength = Math.max(...data.map(row => (row[i] ? row[i].toString().length : 10)));
                    return { wch: Math.min(Math.max(maxLength + 2, 10), 50) }; // Min 10, Max 50
                });

                ws["!cols"] = colWidths;

                let wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

                let buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

                // 🔹 File Path Handling: msg.filepath > UI-configured path > Default path
                let filePath = msg.filepath || config.filepath || path.join(__dirname, "output.xlsx");

                // 🔹 Notify user if msg.filepath is overriding the UI setting
                if (msg.filepath) {
                    node.warn(`📂 File path OVERRIDDEN by msg.filepath: ${msg.filepath}`);
                } else if (config.filepath) {
                    node.log(`📁 Using UI-configured file path: ${config.filepath}`);
                } else {
                    node.warn(`⚠️ No file path set! Using default: ${filePath}`);
                }

                // 🔹 Ensure the directory exists
                let dir = path.dirname(filePath);
                if (!fs.existsSync(dir)) {
                    fs.mkdirSync(dir, { recursive: true });
                }

                // 🔹 Save file
                fs.writeFileSync(filePath, buffer);

                // 🔹 Update msg.payload with result
                msg.payload = {
                    success: true,
                    message: `✅ Excel file saved: ${filePath}`,
                    filepath: filePath
                };

                node.send(msg);
            } catch (error) {
                node.error(`❌ Error processing XLSX: ${error.message}`, msg);
            }
        });
    }

    RED.nodes.registerType("xlsx-autofit", XlsxAutofitNode);
};

