from getpass import getpass

from sqlalchemy.exc import SQLAlchemyError

from database.calibration_store import redefinir_senha_administrador


def main() -> None:
    print("Redefinição local de senha da Gestão de Calibração")
    email = input("E-mail do administrador: ").strip()
    senha = getpass("Nova senha (mínimo 12 caracteres): ")
    confirmacao = getpass("Confirme a nova senha: ")
    if senha != confirmacao:
        raise SystemExit("As senhas não coincidem.")
    try:
        redefinir_senha_administrador(email, senha)
    except (ValueError, OSError, SQLAlchemyError) as erro:
        raise SystemExit(f"Não foi possível redefinir a senha: {erro}") from erro
    print("Senha redefinida. Entre no app com a nova senha.")


if __name__ == "__main__":
    main()
