import subprocess
import sys
import os

def install_requirements():
    # Список необходимых пакетов
    requirements_file = os.path.join(os.path.dirname(__file__), 'requirements.txt')

    console_message = """
    ████████╗██████╗░██╗██████╗░░█████╗░███████╗░█████╗░██████╗░
    ╚══██╔══╝██╔══██╗██║██╔══██╗██╔══██╗██╔════╝██╔══██╗██╔══██╗
    ░░░██║░░░██████╔╝██║██████╔╝██║░░██║█████╗░░██║░░██║██████╦╝
    ░░░██║░░░██╔══██╗██║██╔═══╝░██║░░██║██╔══╝░░██║░░██║██╔══██╗
    ░░░██║░░░██║░░██║██║██║░░░░░╚█████╔╝██║░░░░░╚█████╔╝██████╦╝
    ░░░╚═╝░░░╚═╝░░╚═╝╚═╝╚═╝░░░░░░╚════╝░╚═╝░░░░░░╚════╝░╚═════╝░
    
    Установка необходимых зависимостей...
    """
    print(console_message)

    try:
        # Проверяем, установлен ли pip
        subprocess.check_call([sys.executable, '-m', 'pip', '--version'])
    except subprocess.CalledProcessError:
        print("Ошибка: pip не установлен. Пожалуйста, установите pip сначала.")
        return

    # Устанавливаем зависимости из requirements.txt
    try:
        print("Установка зависимостей из requirements.txt...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', requirements_file])
        print("Все зависимости успешно установлены!")
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при установке зависимостей: {e}")
        return

    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    install_requirements()