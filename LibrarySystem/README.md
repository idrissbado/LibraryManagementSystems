# 📚 Library Management System

A modular, extensible Library Management System built with Node.js and ES modules. This project demonstrates clean architecture, design patterns, and best practices for scalable JavaScript applications.

---

## ✨ Features
- **ES Modules** for modern, maintainable code
- **Design Patterns**: Factory, Strategy, Observer
- **Dependency Injection** for testability
- **Interface-like Abstractions** for flexibility
- **Reusable Utility Modules**
- **Easy to Extend** for new item types or policies

---

## 🏗️ Architecture
- `LibrarySystem/`
  - `utils.js` — Utility functions (ID generation, date formatting)
  - `observer.js` — Observer pattern (Observable, Observer)
  - `factory.js` — Factory pattern for item creation (Book, Magazine, etc.)
  - `strategy.js` — Strategy pattern for lending policies
  - `repository.js` — Interface-like repository & in-memory implementation
  - `library.js` — Main Library class (dependency injection, eventing)
  - `index.mjs` — Example usage & entry point

---

## 🚀 Getting Started

1. **Clone the repo**
   ```sh
   git clone https://github.com/idrissbado/LibraryManagementSystems.git
   cd LibraryManagementSystems
   ```
2. **Run the example**
   ```sh
   node LibrarySystem/index.mjs
   ```

---

## 🧩 Design Patterns Used
- **Factory**: Create different library items (Book, Magazine, etc.)
- **Strategy**: Switchable lending policies (Standard, Premium)
- **Observer**: Subscribe to library events (item added, lent, etc.)

---

## 🛠️ Extending the System
- Add new item types: Extend the factory and item classes
- Add new lending strategies: Implement new strategy classes
- Swap repository: Implement and inject a new repository (e.g., database)

---

## 🧪 Testing & Development
- Write tests for each module (suggested: Jest, Mocha)
- Use dependency injection for easy mocking

---

## 📄 License
MIT

---

## 👤 Author
- [Idriss Bado](https://github.com/idrissbado)

---

## 🌟 Star this repo if you find it useful!
