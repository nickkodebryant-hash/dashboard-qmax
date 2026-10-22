from getpass import getpass

from database.calibration_store import provisionar_administrador


def main() -> None:
    print("Provisionamento do primeiro administrador da Gestão de Calibração")
    nome = input("Nome: ").strip()
    email = input("E-mail: ").strip()
    senha = getpass("Senha (mínimo 12 caracteres): ")
    confirmacao = getpass("Confirme a senha: ")
    if senha != confirmacao:
        raise SystemExit("As senhas não coincidem.")
    try:
        provisionar_administrador(nome, email, senha)
    except (ValueError, OSError) as erro:
        raise SystemExit(str(erro)) from erro
    print("Administrador criado com sucesso.")


if __name__ == "__main__":
    main()
