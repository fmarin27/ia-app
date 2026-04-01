from __future__ import annotations

import json
import subprocess
import tempfile
import threading
from dataclasses import asdict, dataclass
from pathlib import Path, PurePosixPath
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
PROFILES_FILE = APP_DIR / "sftp_profiles.json"
KEYS_DIR = APP_DIR / "sftp_keys"
SFTP_EXE = Path(r"C:\WINDOWS\System32\OpenSSH\sftp.exe")
SSH_KEYGEN_EXE = Path(r"C:\WINDOWS\System32\OpenSSH\ssh-keygen.exe")


@dataclass
class SftpProfile:
    name: str = "Office PC"
    host: str = ""
    port: int = 22
    username: str = ""
    identity_file: str = ""
    upload_remote_dir: str = "/Users"
    download_remote_path: str = "/Users"
    local_upload_folder: str = ""
    local_download_folder: str = ""


class SftpTransferApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Fernando SFTP Transfer Tool")
        self.root.geometry("1180x780")

        self.profiles = self.load_profiles()
        self.profile_names = sorted(self.profiles) or ["Office PC"]
        self.is_running = False

        self.profile_var = tk.StringVar(value=self.profile_names[0])
        self.host_var = tk.StringVar()
        self.port_var = tk.StringVar(value="22")
        self.username_var = tk.StringVar()
        self.identity_var = tk.StringVar()
        self.upload_remote_var = tk.StringVar()
        self.download_remote_var = tk.StringVar()
        self.local_upload_var = tk.StringVar()
        self.local_download_var = tk.StringVar()
        self.command_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready. Use key-based auth for the cleanest setup.")

        self._build_ui()
        self.load_profile_into_form(self.profile_var.get())
        self.refresh_command_preview()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        shell = ttk.Frame(self.root, padding=12)
        shell.grid(sticky="nsew")
        shell.columnconfigure(0, weight=3)
        shell.columnconfigure(1, weight=2)
        shell.rowconfigure(0, weight=1)

        left = ttk.Frame(shell)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(1, weight=1)

        right = ttk.Notebook(shell)
        right.grid(row=0, column=1, sticky="nsew")

        ttk.Label(left, text="Saved Profile").grid(row=0, column=0, sticky="w")
        profile_row = ttk.Frame(left)
        profile_row.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        profile_row.columnconfigure(0, weight=1)
        self.profile_combo = ttk.Combobox(
            profile_row,
            textvariable=self.profile_var,
            values=self.profile_names,
            state="readonly",
        )
        self.profile_combo.grid(row=0, column=0, sticky="ew")
        self.profile_combo.bind("<<ComboboxSelected>>", lambda _event: self.on_profile_selected())
        ttk.Button(profile_row, text="New", command=self.new_profile).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(profile_row, text="Save", command=self.save_current_profile).grid(row=0, column=2, padx=(6, 0))

        self._add_entry(left, 1, "Profile Name", self.profile_var)
        self._add_entry(left, 2, "Host / IP", self.host_var)
        self._add_entry(left, 3, "Port", self.port_var)
        self._add_entry(left, 4, "Username", self.username_var)

        ttk.Label(left, text="Private Key").grid(row=5, column=0, sticky="w", pady=(0, 8))
        key_row = ttk.Frame(left)
        key_row.grid(row=5, column=1, sticky="ew", pady=(0, 8))
        key_row.columnconfigure(0, weight=1)
        ttk.Entry(key_row, textvariable=self.identity_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(key_row, text="Browse", command=self.choose_identity_file).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(key_row, text="Generate Key", command=self.generate_keypair).grid(row=0, column=2, padx=(6, 0))

        ttk.Label(left, text="Local Upload Folder").grid(row=6, column=0, sticky="w", pady=(0, 8))
        upload_local_row = ttk.Frame(left)
        upload_local_row.grid(row=6, column=1, sticky="ew", pady=(0, 8))
        upload_local_row.columnconfigure(0, weight=1)
        ttk.Entry(upload_local_row, textvariable=self.local_upload_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(upload_local_row, text="Choose", command=self.choose_upload_folder).grid(row=0, column=1, padx=(6, 0))

        self._add_entry(left, 7, "Remote Upload Dir", self.upload_remote_var)

        ttk.Label(left, text="Local Download Folder").grid(row=8, column=0, sticky="w", pady=(0, 8))
        download_local_row = ttk.Frame(left)
        download_local_row.grid(row=8, column=1, sticky="ew", pady=(0, 8))
        download_local_row.columnconfigure(0, weight=1)
        ttk.Entry(download_local_row, textvariable=self.local_download_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(download_local_row, text="Choose", command=self.choose_download_folder).grid(row=0, column=1, padx=(6, 0))

        self._add_entry(left, 9, "Remote Download Path", self.download_remote_var)

        button_row = ttk.Frame(left)
        button_row.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(8, 10))
        for column in range(3):
            button_row.columnconfigure(column, weight=1)
        ttk.Button(button_row, text="Upload Folder", command=self.start_upload).grid(row=0, column=0, sticky="ew")
        ttk.Button(button_row, text="Download Folder", command=self.start_download).grid(row=0, column=1, padx=8, sticky="ew")
        ttk.Button(button_row, text="Open Key Folder", command=self.open_key_folder).grid(row=0, column=2, sticky="ew")

        preview_frame = ttk.LabelFrame(left, text="Command Preview", padding=8)
        preview_frame.grid(row=11, column=0, columnspan=2, sticky="nsew")
        left.rowconfigure(11, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        ttk.Label(
            preview_frame,
            textvariable=self.command_var,
            wraplength=620,
            justify="left",
        ).grid(row=0, column=0, sticky="nw")

        for variable in (
            self.profile_var,
            self.host_var,
            self.port_var,
            self.username_var,
            self.identity_var,
            self.upload_remote_var,
            self.download_remote_var,
            self.local_upload_var,
            self.local_download_var,
        ):
            variable.trace_add("write", lambda *_args: self.refresh_command_preview())

        log_tab = ttk.Frame(right, padding=10)
        learn_tab = ttk.Frame(right, padding=10)
        right.add(log_tab, text="Transfer Log")
        right.add(learn_tab, text="How It Works")

        log_tab.columnconfigure(0, weight=1)
        log_tab.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_tab, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        learn_tab.columnconfigure(0, weight=1)
        learn_tab.rowconfigure(0, weight=1)
        learn_text = tk.Text(learn_tab, wrap="word")
        learn_text.grid(row=0, column=0, sticky="nsew")
        learn_text.insert("1.0", self.learning_notes())
        learn_text.config(state="disabled")

        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=(12, 6))
        status_bar.grid(row=1, column=0, sticky="ew")

    def _add_entry(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=(0, 8))

    def load_profiles(self) -> dict[str, SftpProfile]:
        if not PROFILES_FILE.exists():
            return {}
        try:
            raw = json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
        profiles: dict[str, SftpProfile] = {}
        for name, data in raw.items():
            profiles[name] = SftpProfile(**data)
        return profiles

    def save_profiles(self) -> None:
        payload = {name: asdict(profile) for name, profile in self.profiles.items()}
        PROFILES_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def current_profile_from_form(self) -> SftpProfile:
        port_text = self.port_var.get().strip() or "22"
        try:
            port = int(port_text)
        except ValueError:
            raise ValueError("Port must be a whole number.")
        return SftpProfile(
            name=self.profile_var.get().strip() or "Office PC",
            host=self.host_var.get().strip(),
            port=port,
            username=self.username_var.get().strip(),
            identity_file=self.identity_var.get().strip(),
            upload_remote_dir=self.upload_remote_var.get().strip(),
            download_remote_path=self.download_remote_var.get().strip(),
            local_upload_folder=self.local_upload_var.get().strip(),
            local_download_folder=self.local_download_var.get().strip(),
        )

    def load_profile_into_form(self, profile_name: str) -> None:
        profile = self.profiles.get(profile_name, SftpProfile(name=profile_name))
        self.profile_var.set(profile.name)
        self.host_var.set(profile.host)
        self.port_var.set(str(profile.port))
        self.username_var.set(profile.username)
        self.identity_var.set(profile.identity_file)
        self.upload_remote_var.set(profile.upload_remote_dir)
        self.download_remote_var.set(profile.download_remote_path)
        self.local_upload_var.set(profile.local_upload_folder)
        self.local_download_var.set(profile.local_download_folder)

    def on_profile_selected(self) -> None:
        self.load_profile_into_form(self.profile_var.get())

    def new_profile(self) -> None:
        base_name = "Office PC"
        suffix = 1
        candidate = base_name
        while candidate in self.profiles:
            suffix += 1
            candidate = f"{base_name} {suffix}"
        self.load_profile_into_form(candidate)
        self.status_var.set("New profile ready. Fill in the office PC details and save it.")

    def save_current_profile(self) -> None:
        try:
            profile = self.current_profile_from_form()
        except ValueError as exc:
            messagebox.showerror("Invalid Profile", str(exc))
            return
        self.profiles[profile.name] = profile
        self.profile_names = sorted(self.profiles)
        self.profile_combo["values"] = self.profile_names
        self.save_profiles()
        self.profile_var.set(profile.name)
        self.status_var.set(f"Saved profile: {profile.name}")
        self.log(f"Saved profile '{profile.name}' to {PROFILES_FILE.name}.")

    def choose_identity_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose private key",
            initialdir=str(KEYS_DIR if KEYS_DIR.exists() else APP_DIR),
        )
        if path:
            self.identity_var.set(path)

    def choose_upload_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose local folder to upload")
        if path:
            self.local_upload_var.set(path)

    def choose_download_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose local folder to receive downloads")
        if path:
            self.local_download_var.set(path)

    def open_key_folder(self) -> None:
        KEYS_DIR.mkdir(exist_ok=True)
        subprocess.Popen(["explorer.exe", str(KEYS_DIR)])

    def refresh_command_preview(self) -> None:
        host = self.host_var.get().strip() or "host-or-ip"
        username = self.username_var.get().strip() or "username"
        port = self.port_var.get().strip() or "22"
        identity = self.identity_var.get().strip() or r"C:\path\to\id_ed25519"
        self.command_var.set(
            f'{SFTP_EXE.name} -i "{identity}" -P {port} -b <batch-file> '
            f'-oBatchMode=yes -oStrictHostKeyChecking=accept-new {username}@{host}\n\n'
            "This tool writes a tiny batch file with commands like:\n"
            "mkdir /target\ncd /target\nput -R <local-folder>\n"
            "or\nlcd <local-folder>\nget -R <remote-path>"
        )

    def generate_keypair(self) -> None:
        name = self.profile_var.get().strip() or "office_pc"
        safe_name = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in name.lower())
        KEYS_DIR.mkdir(exist_ok=True)
        key_path = KEYS_DIR / f"{safe_name}_id_ed25519"
        if key_path.exists():
            overwrite = messagebox.askyesno("Key Exists", f"{key_path.name} already exists. Create a fresh one anyway?")
            if not overwrite:
                self.identity_var.set(str(key_path))
                return
        comment = f"{self.username_var.get().strip() or 'fernando'}@office-transfer"
        command = [
            str(SSH_KEYGEN_EXE),
            "-t",
            "ed25519",
            "-f",
            str(key_path),
            "-N",
            "",
            "-C",
            comment,
        ]
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode != 0:
            messagebox.showerror("Key Generation Failed", result.stderr.strip() or "ssh-keygen returned an error.")
            return
        self.identity_var.set(str(key_path))
        self.log(f"Created keypair:\n- Private: {key_path}\n- Public: {key_path}.pub")
        self.log("Copy the .pub file contents into the office PC user's authorized_keys file.")
        self.status_var.set("SSH keypair created.")

    def start_upload(self) -> None:
        self.run_transfer("upload")

    def start_download(self) -> None:
        self.run_transfer("download")

    def run_transfer(self, mode: str) -> None:
        if self.is_running:
            messagebox.showinfo("Transfer Running", "Wait for the current transfer to finish.")
            return
        try:
            profile = self.current_profile_from_form()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc))
            return
        validation_error = self.validate_profile(profile, mode)
        if validation_error:
            messagebox.showerror("Missing Info", validation_error)
            return

        self.save_current_profile()
        self.is_running = True
        self.status_var.set(f"Running {mode}...")
        self.log(f"Starting {mode} using profile '{profile.name}'.")
        thread = threading.Thread(target=self._run_transfer_worker, args=(profile, mode), daemon=True)
        thread.start()

    def validate_profile(self, profile: SftpProfile, mode: str) -> str:
        if not SFTP_EXE.exists():
            return f"Could not find {SFTP_EXE}"
        if not profile.host:
            return "Host / IP is required."
        if not profile.username:
            return "Username is required."
        if not profile.identity_file:
            return "Choose or generate a private key first."
        if not Path(profile.identity_file).exists():
            return "The selected private key file does not exist."
        if mode == "upload":
            if not profile.local_upload_folder:
                return "Choose a local upload folder."
            if not Path(profile.local_upload_folder).exists():
                return "The local upload folder does not exist."
            if not profile.upload_remote_dir:
                return "Remote Upload Dir is required."
        else:
            if not profile.local_download_folder:
                return "Choose a local download folder."
            if not Path(profile.local_download_folder).exists():
                return "The local download folder does not exist."
            if not profile.download_remote_path:
                return "Remote Download Path is required."
        return ""

    def _run_transfer_worker(self, profile: SftpProfile, mode: str) -> None:
        batch_path: Path | None = None
        try:
            batch_contents = self.build_batch_contents(profile, mode)
            self.log("SFTP batch commands:\n" + batch_contents)
            with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as batch_file:
                batch_file.write(batch_contents)
                batch_path = Path(batch_file.name)

            command = [
                str(SFTP_EXE),
                "-i",
                profile.identity_file,
                "-P",
                str(profile.port),
                "-b",
                str(batch_path),
                "-oBatchMode=yes",
                "-oStrictHostKeyChecking=accept-new",
                f"{profile.username}@{profile.host}",
            ]
            self.log("Running command:\n" + self.render_command(command))
            result = subprocess.run(command, capture_output=True, text=True)
            combined_output = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
            if combined_output:
                self.log(combined_output)
            if result.returncode != 0:
                raise RuntimeError("SFTP transfer failed. Check the transfer log for details.")
            self.root.after(0, lambda: self.status_var.set(f"{mode.capitalize()} finished successfully."))
            self.root.after(0, lambda: messagebox.showinfo("Transfer Complete", f"{mode.capitalize()} finished successfully."))
        except Exception as exc:
            self.root.after(0, lambda: self.status_var.set(str(exc)))
            self.root.after(0, lambda: messagebox.showerror("Transfer Failed", str(exc)))
        finally:
            if batch_path and batch_path.exists():
                try:
                    batch_path.unlink()
                except OSError:
                    pass
            self.root.after(0, self.finish_transfer)

    def finish_transfer(self) -> None:
        self.is_running = False

    def build_batch_contents(self, profile: SftpProfile, mode: str) -> str:
        if mode == "upload":
            remote_dir = self.normalize_remote(profile.upload_remote_dir)
            local_folder = Path(profile.local_upload_folder)
            commands = [
                *self.mkdir_commands(remote_dir),
                f'cd "{remote_dir}"',
                f'put -R "{self.to_sftp_local_path(local_folder)}"',
            ]
        else:
            remote_path = self.normalize_remote(profile.download_remote_path)
            local_dir = Path(profile.local_download_folder)
            commands = [
                f'lcd "{self.to_sftp_local_path(local_dir)}"',
                f'get -R "{remote_path}"',
            ]
        return "\n".join(commands) + "\n"

    def normalize_remote(self, remote_path: str) -> str:
        text = remote_path.strip().replace("\\", "/")
        if not text.startswith("/"):
            text = "/" + text
        return str(PurePosixPath(text))

    def mkdir_commands(self, remote_dir: str) -> list[str]:
        path = PurePosixPath(remote_dir)
        commands: list[str] = []
        current = PurePosixPath("/")
        for part in path.parts[1:]:
            current = current / part
            commands.append(f'mkdir "{current.as_posix()}"')
        return commands

    def to_sftp_local_path(self, path: Path) -> str:
        return path.resolve().as_posix()

    def render_command(self, command: list[str]) -> str:
        return " ".join(f'"{item}"' if " " in item else item for item in command)

    def log(self, message: str) -> None:
        def write() -> None:
            self.log_text.config(state="normal")
            self.log_text.insert("end", message.rstrip() + "\n\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")

        self.root.after(0, write)

    def learning_notes(self) -> str:
        return (
            "Why this works\n"
            "\n"
            "1. SFTP is the file-transfer subsystem built on top of SSH.\n"
            "2. Your Python app does not have to implement the protocol itself.\n"
            "3. Instead, this tool shells out to Windows' built-in sftp.exe client.\n"
            "4. The client authenticates with an SSH key and then runs a tiny batch file of SFTP commands.\n"
            "\n"
            "The real moving parts\n"
            "\n"
            "- ssh-keygen.exe creates a private key and a matching public key.\n"
            "- The office PC stores the public key in authorized_keys.\n"
            "- sftp.exe connects to username@host on port 22.\n"
            "- Batch commands like mkdir, cd, put -R, lcd, and get -R do the actual transfer.\n"
            "\n"
            "What to reuse in claims_manager.py later\n"
            "\n"
            "- Profile storage: save host, port, username, remote paths, and key path as JSON.\n"
            "- Validation: check local paths, key file, and required connection settings before a transfer.\n"
            "- Execution: write a temporary batch file, call sftp.exe with subprocess.run, and capture stdout/stderr.\n"
            "- UX: keep the transfer on a background thread and stream the log back into the UI.\n"
            "\n"
            "Tomorrow's office setup\n"
            "\n"
            "1. On the new office PC, install or enable OpenSSH Server.\n"
            "2. Start the sshd service there.\n"
            "3. Put this tool's .pub key into the office user's authorized_keys file.\n"
            "4. Find that PC's IP address.\n"
            "5. Enter the host, username, key, and remote folder here, then upload.\n"
            "\n"
            "Important note\n"
            "\n"
            "This utility is intentionally key-based. Plain password automation with the stock Windows sftp.exe client is clunky, while key-based auth is cleaner, safer, and much easier to embed in your claims manager app."
        )


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = SftpTransferApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
