# Core

## Visão Geral

 `Core` é projetado para facilitar tarefas de automação envolvendo gerenciamento de janelas e interações baseadas em imagens em um sistema operacional Windows. Esta classe utiliza várias bibliotecas, como `pyautogui`, `pygetwindow`, `pywinauto` e `ctypes`, para interagir com a interface gráfica do Windows e realizar tarefas automatizadas.

## Atributos

- `folder_path` (str): O caminho do diretório onde o script atual está localizado.
- `png` (dict): Um dicionário contendo nomes de arquivos de imagem associados às suas chaves.

## Métodos

`get_active_window_names() -> list[str]`
Retorna uma lista de todos os títulos de janelas ativas no momento.

`get_active_window_name() -> str`
Retorna o título da janela atualmente ativa.

`open_window(window_title: str) -> None`
Foca na janela com o título especificado se ela estiver atualmente ativa.

`open_software(programdir: str) -> None`
Abre um aplicativo de software localizado no diretório especificado.

`wait_until_window_is_open(window: str, timeout: int = 20) -> None`
Espera até que a janela especificada se torne ativa ou lança um TimeoutError.

`execute_image_based_write( auth_img_array: list[str], auth_str_array: list[str], diff: tuple[tuple[int, int], tuple[int, int], tuple[int, int]] = ((0, 0), (0, 0), (0, 0)) ) -> None`
Realiza interações baseadas em imagens com a tela, como cliques e digitação, com base em imagens e strings fornecidas.

`search_open_and_auth(
    window_title: str,
    has_login: bool = False,
    login_page: str = "",
    user: str = "",
    password: str = "",
    diff: tuple[tuple[int, int], tuple[int, int], tuple[int, int]] = ((0, 0), (0, 0), (0, 0)),
    auth_image_path_array: list[str] | tuple[str, str, str] = ("", "", "")
) -> None`
Verifica se a janela especificada está aberta e tenta abrir o software se não estiver. Se o login for necessário, realiza a autenticação.

`find_image_path(dict_key: str) -> str`
Encontra o caminho para um arquivo de imagem no diretório de trabalho atual com base em uma chave do dicionário de imagens.

`find_img_and_click(
    dict_key: str,
    difx: int = 0,
    dify: int = 0,
    delay: float = 0.2
) -> None`
Encontra uma imagem na tela e realiza um clique em sua localização com offsets e atraso opcionais.

`try_click_one_or_more(*dict_key_tuple: str) -> None`
Tenta encontrar e clicar em cada imagem especificada na tela.

`wait_until_is_on_screen(
    key: str,
    timeout: int = 200,
    poll_interval: int = 2
) -> None`
Espera até que uma imagem especificada apareça na tela ou expira após o tempo dado.

`alert_window(message: str, title: str) -> None`
Exibe uma caixa de mensagem do Windows com a mensagem e o título fornecidos.

## Exemplos de Uso

```python
# Inicializa a classe Core
core = Core()

# Abre uma janela específica
core.open_window("Sem título - Bloco de Notas")

# Abre um aplicativo de software
core.open_software("C:\\Program Files\\SomeSoftware\\application.exe")

# Espera até que uma janela se torne ativa
core.wait_until_window_is_open("Título da Janela")

# Realiza escrita baseada em imagens (clique e digitação)
core.execute_image_based_write(
    auth_img_array=["login_button.png", "username_field.png", "password_field.png"],
    auth_str_array=["", "meu_usuario", "minha_senha"]
)

# Busca, abre e autentica se necessário
core.search_open_and_auth(
    window_title="Janela do Aplicativo",
    has_login=True,
    login_page="Página de Login",
    user="meu_usuario",
    password="minha_senha",
    auth_image_path_array=("login_button.png", "username_field.png", "password_field.png")
)

# Encontra uma imagem e clica nela
core.find_img_and_click("submit_button")

# Tenta clicar em uma ou mais imagens
core.try_click_one_or_more("button1", "button2", "button3")

# Espera até que uma imagem apareça na tela
core.wait_until_is_on_screen("loading_spinner")

# Exibe uma janela de alerta
core.alert_window("Operação concluída com sucesso.", "Sucesso")
```
## Dependências

Certifique-se de que as seguintes bibliotecas estão instaladas:

- `ctypes` (parte da biblioteca padrão do Python)
- `pyautogui`
- `pygetwindow`
- `pywinauto`

Além disso, o módulo `files` deve conter um dicionário `png` com o nome e uma key para as imagens necessárias que estão na pasta images.

No arquivo `files.py`, defina o dicionário `png` com os nomes dos arquivos de imagem, por exemplo:

```python
png = {
    "login_button": "login_button.png",
    "username_field": "username_field.png",
    "password_field": "password_field.png"
}
```
## Estrutura de Diretórios

Certifique-se de que o diretório de imagens está corretamente configurado e contém os arquivos de imagem necessários. A estrutura de diretórios deve ser semelhante a:

/path/to/your/project/ ├── your_script.py ├── images/ │ ├── login_button.png │ ├── username_field.png │ └── password_field.png └── files.py
