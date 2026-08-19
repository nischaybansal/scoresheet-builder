<div align="center">

# Scoresheet Builder

**SHL · Professional Services**

![Access](https://img.shields.io/badge/access-internal-1B3A57?style=flat-square)
![Python](https://img.shields.io/badge/python-0E7C86?style=flat-square&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/streamlit-0E7C86?style=flat-square&logo=streamlit&logoColor=white)
![Status](https://img.shields.io/badge/status-active-4B7F52?style=flat-square)

</div>

---

> [!IMPORTANT]
> **Internal repository.** This codebase supports an internal SHL Professional Services
> workflow. It is not a product, carries no external support, and is not intended for
> distribution or use outside the team.

<br>

## Repository

| Path | Role |
| :--- | :--- |
| `streamlit_app.py` | Interface layer |
| `generator.py` | Processing engine |
| `.streamlit/config.toml` | Presentation settings |
| `requirements.txt` | Dependency manifest |

The two modules are deliberately separate: the engine holds no interface code, and the
interface holds no output logic. Either can be replaced without disturbing the other.

<br>

## Access

Environment setup, deployment and day-to-day use are handled internally and are not
documented here. Requests for access should go through the repository owner.

<br>

## Contributing

Changes are made by arrangement with the maintainer. Please raise an issue describing the
intended change before opening a pull request, and keep the separation between the engine
and the interface intact.

<br>

---

<div align="center">
<sub>

**Confidential.** © SHL. Internal use only — not for redistribution.

</sub>
</div>
