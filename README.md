# 🔍 EAN Code Validator & Highlighter (Excel VBA Macro)

This Excel VBA macro automates the process of validating **EAN-13 product codes** in a selected column and highlights any **invalid entries** with a blue background and white font. It ensures consistency and correctness of EAN codes by applying validation rules, including **pattern checks and check digit verification**.

---

## 📌 Features

- 📦 **Validates EAN-13 codes**:
  - Must be 13-digit numeric
  - Must not start with `2`
  - Must not contain disallowed patterns (e.g., `000`, `00000`)
  - Must pass **EAN-13 check digit** algorithm

- 🔄 **Automatically pads shorter EANs** with leading zeros (up to 13 digits)

- 📊 **Interactive column selection** via user input

- 🎨 **Highlights invalid EANs** with:
  - Blue cell background (`#00B0F0`)
  - White text for better visibility

---

## 🛠️ How It Works

1. **Prompts** the user to enter a column letter (e.g., `A`, `AA`)
2. **Scans** the column from row 2 downward (assumes row 1 is header)
3. For each cell:
   - Pads the value with zeros if it's less than 13 digits
   - Validates the final EAN using strict rules and checksum
   - Highlights the cell if the value is invalid

---

## 📁 Project Structure

