import subprocess
import json
import shutil
import sys
import re
import os
import winreg

def find_office_path(program_name: str) -> str:
    """
    Finds the installation path of an Office program by searching the Windows Registry.

    Args:
        program_name (str): The program name to search for (e.g., 'winword', 'excel', 'powerpnt').

    Returns:
        str: The full path to the executable, or an error message if not found.
    """
    # Map friendly names to known Office executable names
    office_programs = {
        "word": "WINWORD.EXE",
        "excel": "EXCEL.EXE",
        "powerpoint": "POWERPNT.EXE",
        "outlook": "OUTLOOK.EXE",
        "access": "MSACCESS.EXE",
        "publisher": "MSPUB.EXE"
    }
    
    # Get the actual executable name from the mapping
    executable_name = office_programs.get(program_name.lower(), program_name)
    
    try:
        # Open the App Paths key in the registry
        key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executable_name}"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            # Read the default value which contains the executable path
            program_path, _ = winreg.QueryValueEx(key, None)
            return program_path
    except FileNotFoundError:
        return f"Error: {program_name} not found in the registry."
    except Exception as e:
        return f"Error while accessing registry: {str(e)}"



def find_program_path(program_name: str) -> str:
    """
    Finds the executable path of a program by mapping common names to executables.
    """
    # Map the program name to an executable
    executable_name = program_name.lower()
    # Try to locate the executable using shutil.which()
    path = shutil.which(executable_name)
    if path:
        return path

    # # Try to locate Microsoft Word in the registry
    # if program_name.lower() in {"microsoft word", "word"}:
    #     try:
    #         key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE")
    #         path, _ = winreg.QueryValueEx(key, None)
    #         return path
    #     except FileNotFoundError:
    #         return None

    return None

def open_program(program_name: str) -> str:
    """
    Opens a program by name, resolving its path first.
    """
    path = find_program_path(program_name)
    if not path: path = find_office_path(program_name)
    if path:
        try:
            print(f"Path for {program_name}: {path}")
            subprocess.Popen(f'"{path}"', shell=True)
            return f"Program '{program_name}' opened successfully."
        except Exception as e:
            return f"Error while opening program: {str(e)}"
    else:
        try:
            # 启动程序
            subprocess.Popen(program_name, shell=True)
            os.startfile(program_name)
            return f"Program '{program_name}' opened successfully."
        except FileNotFoundError:
            return f"Error: Program '{program_name}' not found."
        except Exception as e:
            return f"Error while opening program: {str(e)}"

        return f"Error: Program '{program_name}' not found."

def construct_command(action: str, details: dict) -> str:
    """
    Constructs a system command based on the action and details provided by the model.
    """
    if action == "open_program":
        program = details.get("program", "")
        if not program:
            return "Error: No program specified to open."
        return open_program(program)

    elif action == "list_directory":
        path = details.get("path", ".")
        return f"dir {path}" if sys.platform.startswith("win") else f"ls {path}"

    elif action == "show_path":
        return "echo %cd%" if sys.platform.startswith("win") else "pwd"

    return "Error: Unsupported action."

def run_command(command: str) -> str:
    """
    Executes the given system command.
    """
    if command.startswith("Error"):
        return command
    try:
        result = subprocess.run(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode == 0:
            return result.stdout
        else:
            return f"Error: {result.stderr}"
    except Exception as e:
        return f"Error while executing command: {str(e)}"

def query_model(prompt: str) -> str:
    """
    Sends a prompt to the Ollama model and retrieves the response.
    """
    try:
        result = subprocess.run(
            ["ollama", "run", "llama3.2"],
            input=prompt,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode == 0:
            raw_output = result.stdout.strip()
            # Extract JSON using regex
            match = re.search(r"\{.*\}", raw_output, re.DOTALL)
            if match:
                return match.group(0)
            else:
                return "Error: No valid JSON found in model output."
        else:
            return f"Error from model: {result.stderr}"
    except Exception as e:
        return f"Error while querying model: {str(e)}"

def main():
    """
    Main loop: interact with the user, query the model, and execute commands.
    """
    print("欢迎使用智能助手！输入您的请求，或输入 'exit' 退出程序。")
    system_prompt = """你是一个助手。当用户提出请求时，请输出以下 JSON 格式的内容：
{
  "action": "具体动作，比如 open_program, list_directory, show_path",
  "program": "程序名称，例如 Microsoft Word, Notepad, Chrome（仅在 open_program 下需要,如果用户输入的是简历文字尽量补齐全部内容， 比如用户给出edge,最好补全 Microsoft Edge.）"
}
"""

    while True:
        user_input = input("\n用户请求: ").strip()
        if user_input.lower() in {"exit", "quit"}:
            print("再见！")
            break

        full_prompt = system_prompt + f"\n用户请求: {user_input}"
        print("\n正在查询模型...")
        model_output = query_model(full_prompt)
        print("模型输出:", model_output)

        if "Error" in model_output:
            print(model_output)
            continue

        try:
            data = json.loads(model_output)
            action = data.get("action")
            if not action:
                print("Error: No action provided by the model.")
                continue

            command = construct_command(action, data)
            print("生成的命令:", command)

            # Execute only if it's not an "open_program" action
            if action != "open_program":
                result = run_command(command)
                print("执行结果:\n", result)
            else:
                print("执行结果:\n", command)

        except json.JSONDecodeError:
            print("Error: Invalid JSON output from model.")

if __name__ == "__main__":
    main()

