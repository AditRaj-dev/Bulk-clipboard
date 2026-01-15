import subprocess
import config
from logger import log

# fallbacks in case config doesn't define templates
AHK_TEMPLATE = getattr(config, "AHK_TEMPLATE", "template.ahk")
AHK_GENERATED = getattr(config, "AHK_GENERATED", "generated.ahk")


def generate_and_run(loop_count: int):
    try:
        with open(AHK_TEMPLATE, "r", encoding="utf-8") as f:
            script = f.read()
    except Exception as e:
        log(f"ERROR reading AHK template: {e}")
        return

    script = script.replace("{{LOOP_COUNT}}", str(loop_count))

    try:
        with open(AHK_GENERATED, "w", encoding="utf-8") as f:
            f.write(script)
    except Exception as e:
        log(f"ERROR writing generated AHK: {e}")
        return

    log(f"AHK script generated with loop count = {loop_count}")
    try:
        subprocess.Popen(["autohotkey", AHK_GENERATED], shell=True)
        log("AHK script launched")
    except Exception as e:
        log(f"ERROR launching AHK: {e}")


def generate_and_run_with_files(loop_count: int, template_path: str, generated_path: str):
    try:
        with open(template_path, "r", encoding="utf-8") as f:
            script = f.read()
    except Exception as e:
        log(f"ERROR reading AHK template '{template_path}': {e}")
        return

    script = script.replace("{{LOOP_COUNT}}", str(loop_count))

    try:
        with open(generated_path, "w", encoding="utf-8") as f:
            f.write(script)
    except Exception as e:
        log(f"ERROR writing generated AHK '{generated_path}': {e}")
        return

    log(f"AHK script '{generated_path}' generated with loop count = {loop_count}")
    try:
        subprocess.Popen(["autohotkey", generated_path], shell=True)
        log(f"AHK script '{generated_path}' launched")
    except Exception as e:
        log(f"ERROR launching AHK '{generated_path}': {e}")


def run_ahk_file(script_path: str):
    try:
        subprocess.Popen(["autohotkey", script_path], shell=True)
        log(f"AHK file launched: {script_path}")
    except Exception as e:
        log(f"ERROR launching AHK file '{script_path}': {e}")
