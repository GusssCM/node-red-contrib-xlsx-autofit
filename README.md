# node-red-contrib-xlsx-autofit

📊 A **Node-RED node** that converts **JSON data to Excel (XLSX)** with **auto-adjusted column widths** and **dynamic file saving**.

### 🚀 Features
✔️ **Convert JSON to Excel (XLSX)**  
✔️ **Autofit column width for better readability**  
✔️ **Dynamic file path support (`msg.filepath`)**  
✔️ **Error handling and debug logging**  
✔️ **Works like `node-red-contrib-excel` but with improvements**  

## 📥 Installation

### 🔹 Install from npm (Recommended)
To install this node in your Node-RED environment:
```bash
cd ~/.node-red
npm install node-red-contrib-xlsx-autofit
node-red-restart

cd ~/.node-red
git clone https://github.com/YOUR_GITHUB_USERNAME/node-red-contrib-xlsx-autofit.git
cd node-red-contrib-xlsx-autofit
npm install
node-red-restart**
```

## 📌 File Path Handling
This node allows **dynamic file saving**. The output **XLSX file path** is determined as follows:

1️⃣ **If `msg.filepath` is set in a previous function node**, that path is used.  
2️⃣ **If `msg.filepath` is NOT set**, the path from the **node UI settings** is used.  
3️⃣ **If neither is set**, the default path is:  
   ```bash
   /home/user/output.xlsx
```
## ❓ Troubleshooting

### 🔴 My file isn’t saving!
✅ **Check if `msg.filepath` is set** (it overrides the UI setting).  
✅ **Ensure the folder exists and is writable.**  

### 🔴 My columns aren’t fitting correctly
✅ Make sure **each object has the same keys** in JSON format.  
✅ If you use **array input**, ensure **headers are included** in the first row.  

## 🛠️ Contributing
Want to improve this node? Here's how:

1️⃣ **Fork this repository** on GitHub  
2️⃣ **Make changes** and test in Node-RED  
3️⃣ **Submit a pull request**  

🔹 **Found a bug?** Report it in the [Issues](https://github.com/gushed/node-red-contrib-xlsx-autofit/issues).  
