# Proyecto Despachos – Automatización de carga en SIGHEOS

Este módulo automatiza la carga de insumos médicos en el sistema SIGHEOS a partir de una planilla Excel predefinida. El objetivo es eliminar la necesidad de ingresar manualmente cada ítem, evitando errores comunes y ahorrando tiempo considerable.

---

## 🧠 ¿Qué hace?

1. El usuario registra en un Excel todos los insumos que desea cargar.
2. El bot (programado en Python con Selenium) **simula el accionar humano** sobre el sistema SIGHEOS:
   - Inicia sesión.
   - Navega al módulo de Despachos.
   - Itera sobre los ítems del Excel.
   - Carga los insumos uno por uno, evitando dobles clics accidentales (como presionar dos veces “Agregar Insumo”).

3. El bot **previene errores humanos comunes**, como duplicaciones o saltos de renglón.
4. Acelera enormemente el proceso: lo que antes llevaba 20 a 30 minutos (o más), ahora se realiza en aproximadamente 1 minuto.

---

## 🧾 Manejo de stock insuficiente

Cuando SIGHEOS detecta que un insumo tiene **stock insuficiente**, el sistema anula esa línea de carga. En lugar de perder ese registro, nuestro bot:

- **Guarda el intento en un archivo `remanentes.json`**, con toda la info necesaria (insumo, cantidad faltante, fecha, etc.).
- Luego, otro módulo (`Remanentes.py`) **reintenta despachar esos ítems acumulados**, intentando completar la operación cuando haya stock disponible.

---

## 🔄 Flujo General

```plaintext
Excel → BotDespacho.py (Selenium) → SIGHEOS
                          ↓
        Stock insuficiente → remanentes.json
                          ↓
                Remanentes.py → SIGHEOS (reintento futuro)
```

---

## 🗃 Archivos importantes

- `BotDespacho.py` → carga todos los ítems del Excel en SIGHEOS.
- `remanentes.json` → archivo acumulativo de ítems no despachados por falta de stock.
- `Remanentes.py` → despacha esos ítems cuando el stock se repone.
- `config.xlsx` → plantilla con insumos a despachar.

---

## ⚠️ Requisitos

- Python 3.x
- Selenium
- ChromeDriver
- Planilla Excel con los insumos
- Acceso a SIGHEOS (usuario con permisos de despacho)

---

## ✨ Beneficios

- Evita errores humanos como dobles clics.
- Mejora la trazabilidad de lo que no se pudo despachar.
- Permite visualizar y auditar fácilmente los remanentes.
- Acelera enormemente la tarea de carga repetitiva.

---
