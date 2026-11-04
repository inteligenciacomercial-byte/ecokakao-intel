# ECO-KAKAO Market Intelligence

Web de inteligencia comercial de cacao.

- `index.html`: la aplicación.
- `imports.json`: importaciones oficiales por país (HS 1801–1806). Lo rellena el robot cada mes.
- `update_comtrade.py`: el robot que consulta UN Comtrade.
- `.github/workflows/`: programa el robot (día 3 de cada mes) y publica la web.

La clave de UN Comtrade se guarda como secret `COMTRADE_KEY` y nunca en el código.

**Importante:** este repositorio es público. No subir aquí datos de contactos, precios ni información interna de ECO-KAKAO.
