# Contributing to ZapSender

First of all, thank you for your interest in contributing to ZapSender! 

## ⚖️ Legal Agreement (Mandatory)

By submitting a Pull Request, you acknowledge and agree to the following:
1. **License:** Your contribution will be licensed under the **GNU AGPLv3**.
2. **Commercial Rights:** You grant the Project Founder (João Henrique Alves Ferreira) the right to include, modify, and distribute your contribution in proprietary/commercial versions of the project (e.g., the GUI/Premium version).
3. **CLA:** You must sign our Contributor License Agreement via the **CLA Assistant** bot before your code can be merged.

## 💻 Environment & Platform Support

**Important:** ZapSender was originally developed for **Windows**. 
* **Primary Support:** Windows 10/11.
* **Experimental Support:** Linux. 

If you are contributing, please prioritize testing your changes on **Windows** to ensure the widest compatibility for our end-users. If your fix is Linux-specific, please mention it clearly in your Pull Request.

## 🛠️ Development Guidelines

### 1. Ethical Automation
We are committed to "White Hat" automation. We will **reject** any PR that:
* Attempts to bypass WhatsApp's security in a malicious way.
* Removes mandatory wait times (*cooldowns*) designed to prevent spam.
* Implements contact scraping or data harvesting features.

### 2. Code Standards
* Use **Black** or follow **PEP 8** for code formatting.
* Use `pathlib` or `os.path.join` for file paths to maintain cross-platform compatibility.
* Ensure all browser driver dependencies are handled gracefully (using `webdriver-manager` is preferred).

## 🚀 How to Contribute

1. **Fork** the repository.
2. Create a **Feature Branch** (`git checkout -b feature/MyNewFeature`).
3. **Commit** your changes with clear messages.
4. **Push** to your branch.
5. Open a **Pull Request** and wait for the CLA bot instructions.

---
*Questions or Partnership inquiries? Contact: joaohenrique.jh103@protonmail.com*
