import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import time
from itertools import cycle
from PIL import Image, ImageTk
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.tenant.administration.tenant import Tenant

# Branding and Colors
BRANDING_COLORS = {
    "background": "#F0ECE3",
    "text": "#0B0B0B",
    "button": "#FBC53C",
    "button_text": "#0B0B0B",
    "progress": "#FFC123",
    "highlight": "#FF931F",
    "banner_background": "#FFC123",
    "banner_text": "#0B0B0B",
}
FONT_PRIMARY = ("LT Wave Text Black", 14)
FONT_TITLE = ("LT Wave Text Black", 20, "bold")
FONT_SECONDARY = ("Montserrat Bold", 16)

BRANDING_MESSAGES = cycle([

"Cybersecurity\n" 
"🔒Defend your digital world with Dapango Cybersecurity\nProactive, reliable, and always vigilant!",

"Compliance\n"
"🛡️Achieve peace of mind with Dapango Compliance\nYour partner in staying ahead of regulations!",

"Disaster Recovery\n"
"🔄Bounce back stronger with Dapango Disaster Recovery\nYour safety net for the unexpected!",

"Integrated IT and Cloud Management\n"
"💻🌩️Elevate your operations with Dapango IT & Cloud Management\nSeamless, scalable, and smart!",

"Virtual CIO (vCIO) Services\n"
"🧠Strategize your success with Dapango vCIO\nVisionary guidance for technology-driven growth!",

"Business Intelligence\n"
"📊Turn data into decisions with Dapango Business Intelligence\nInsights that power your future!",
])

SKIP_LIBRARIES = {"Forms", "_SiteTemplates", "Plantillas_de_formulario", "Activos_del_sitio", "Biblioteca_de_estilos"}

class SharePointBackupApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dapango Technologies SharePoint Backup Utility")
        self.geometry("950x750")
        self.configure(bg=BRANDING_COLORS["background"])

        # Backup State Variables
        self.admin_url = tk.StringVar()
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.backup_folder = tk.StringVar()
        self.backup_mode = tk.StringVar(value="full")
        self.cancel_requested = threading.Event()
        self.is_backup_running = False

        self.current_frame = None
        self.banner_label = None
        self.show_mode_selection()

    def switch_frame(self, frame_class):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = frame_class(self)
        self.current_frame.pack(fill="both", expand=True)

    def show_banner(self, parent):
        self.banner_label = tk.Label(
            parent,
            text=next(BRANDING_MESSAGES),
            font=FONT_PRIMARY,
            bg=BRANDING_COLORS["banner_background"],
            fg=BRANDING_COLORS["banner_text"],
            wraplength=800,
            justify="left",
            padx=10,
            pady=10,
        )
        self.banner_label.pack(pady=10, fill="x")
        self.update_marketing_banner()

        # Adding an image banner
        banner_image_frame = tk.Frame(parent, bg=BRANDING_COLORS["background"])
        banner_image_frame.pack(pady=10)

        try:
            image = Image.open("ondo Teams.jpg")  # Update with actual path
            image = image.resize((800, 150), Image.ANTIALIAS)
            image_banner = ImageTk.PhotoImage(image)
            banner_label = tk.Label(
                banner_image_frame, image=image_banner, bg=BRANDING_COLORS["background"]
            )
            banner_label.image = image_banner  # Prevent garbage collection
            banner_label.pack()
        except Exception as e:
            print(f"Error loading banner image: {e}")

    def update_marketing_banner(self):
        if self.banner_label:
            self.banner_label.config(text=next(BRANDING_MESSAGES))
            self.after(5000, self.update_marketing_banner)

    def show_mode_selection(self):
        self.switch_frame(ModeSelectionFrame)

    def show_backup_configuration(self):
        self.switch_frame(BackupConfigurationFrame)

    def show_backup_progress(self):
        if not self.is_backup_running:
            self.is_backup_running = True
        if self.is_backup_running:
            self.switch_frame(BackupProgressFrame)
        else:
            messagebox.showinfo("No Backup Running", "No ongoing backup to resume.")
            self.show_mode_selection()

class ModeSelectionFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BRANDING_COLORS["background"])
        master.show_banner(self)

        tk.Label(
            self,
            text="Welcome to Dapango Technologies \nSharePoint Backup Utility",
            font=FONT_TITLE,
            bg=BRANDING_COLORS["background"],
            fg=BRANDING_COLORS["text"],
        ).pack(pady=20)

        button_frame = tk.Frame(self, bg=BRANDING_COLORS["background"])
        button_frame.pack(pady=20)

        tk.Button(
            button_frame,
            text="Start Full Backup",
            command=lambda: self.select_mode("full"),
            bg=BRANDING_COLORS["button"],
            fg=BRANDING_COLORS["button_text"],
            font=FONT_PRIMARY,
            width=20,
            height=2
        ).grid(row=0, column=0, padx=10, pady=10)

        tk.Button(
            button_frame,
            text="Start Update Backup",
            command=lambda: self.select_mode("update"),
            bg=BRANDING_COLORS["progress"],
            fg=BRANDING_COLORS["button_text"],
            font=FONT_PRIMARY,
            width=20,
            height=2
        ).grid(row=0, column=1, padx=10, pady=10)

    def select_mode(self, mode):
        self.master.backup_mode.set(mode)
        self.master.show_backup_configuration()

class BackupConfigurationFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BRANDING_COLORS["background"])
        master.show_banner(self)

        tk.Label(
            self,
            text="Configure Backup Settings",
            font=FONT_TITLE,
            bg=BRANDING_COLORS["background"],
            fg=BRANDING_COLORS["text"],
        ).pack(pady=20)

        input_frame = tk.Frame(self, bg=BRANDING_COLORS["background"])
        input_frame.pack(pady=10)

        tk.Label(input_frame, text="Admin URL:", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"]).grid(row=0, column=0, sticky="e")
        tk.Entry(input_frame, textvariable=master.admin_url, width=50).grid(row=0, column=1)

        tk.Label(input_frame, text="Username:", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"]).grid(row=1, column=0, sticky="e")
        tk.Entry(input_frame, textvariable=master.username, width=50).grid(row=1, column=1)

        tk.Label(input_frame, text="Password:", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"]).grid(row=2, column=0, sticky="e")
        tk.Entry(input_frame, textvariable=master.password, width=50, show="*").grid(row=2, column=1)

        tk.Label(input_frame, text="Backup Folder:", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"]).grid(row=3, column=0, sticky="e")
        tk.Entry(input_frame, textvariable=master.backup_folder, width=50).grid(row=3, column=1)
        tk.Button(
            input_frame, text="Browse", command=self.browse_folder, bg=BRANDING_COLORS["button"], fg=BRANDING_COLORS["button_text"]
        ).grid(row=3, column=2)

        button_frame = tk.Frame(self, bg=BRANDING_COLORS["background"])
        button_frame.pack(pady=20)

        tk.Button(
            button_frame,
            text="Start Backup",
            command=master.show_backup_progress,
            bg=BRANDING_COLORS["button"],
            fg=BRANDING_COLORS["button_text"],
            font=FONT_PRIMARY
        ).grid(row=0, column=0, padx=10)

        tk.Button(
            button_frame,
            text="Back",
            command=master.show_mode_selection,
            bg=BRANDING_COLORS["highlight"],
            fg=BRANDING_COLORS["button_text"],
            font=FONT_PRIMARY
        ).grid(row=0, column=1, padx=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.master.backup_folder.set(folder)

class BackupProgressFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BRANDING_COLORS["background"])
        master.show_banner(self)

        tk.Label(
            self, text="Backup Progress", font=FONT_TITLE, bg=BRANDING_COLORS["background"], fg=BRANDING_COLORS["text"]
        ).pack(pady=20)

        self.global_progress = ttk.Progressbar(self, orient="horizontal", length=800, mode="determinate")
        self.global_progress.pack(pady=10)

        self.global_progress_label = tk.Label(self, text="0/0 Sites Processed", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"])
        self.global_progress_label.pack()

        self.folder_progress = ttk.Progressbar(self, orient="horizontal", length=800, mode="determinate")
        self.folder_progress.pack(pady=10)

        self.folder_progress_label = tk.Label(self, text="0/0 Folders Processed", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"])
        self.folder_progress_label.pack()

        self.site_progress = ttk.Progressbar(self, orient="horizontal", length=800, mode="determinate")
        self.site_progress.pack(pady=10)

        self.site_progress_label = tk.Label(self, text="0/0 Files Processed", font=FONT_PRIMARY, bg=BRANDING_COLORS["background"])
        self.site_progress_label.pack()

        self.log_text = tk.Text(self, width=80, height=15, state="disabled", bg="white", font=FONT_PRIMARY)
        self.log_text.pack(pady=20)

        button_frame = tk.Frame(self, bg=BRANDING_COLORS["background"])
        button_frame.pack(pady=20)

        tk.Button(
            button_frame,
            text="Cancel",
            command=self.cancel_backup,
            bg="firebrick",
            fg="white",
            font=FONT_PRIMARY
        ).grid(row=0, column=0, padx=10)

        tk.Button(
            button_frame,
            text="Return to Main Menu",
            command=master.show_mode_selection,
            bg=BRANDING_COLORS["highlight"],
            fg=BRANDING_COLORS["button_text"],
            font=FONT_PRIMARY
        ).grid(row=0, column=1, padx=10)

        threading.Thread(target=self.perform_backup, daemon=True).start()

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def cancel_backup(self):
        self.master.cancel_requested.set()
        self.master.is_backup_running = False
        self.log("Backup canceled by user.")

    def perform_backup(self):
        try:
            self.log("Connecting to SharePoint...")
            admin_ctx = ClientContext(self.master.admin_url.get()).with_credentials(
                UserCredential(self.master.username.get(), self.master.password.get())
            )

            tenant = Tenant(admin_ctx)
            site_props = tenant.get_site_properties_from_sharepoint_by_filters("", 0, True)
            admin_ctx.execute_query()

            total_sites = len(site_props)
            self.global_progress["maximum"] = total_sites

            if total_sites == 0:
                self.log("No sites found. Backup complete.")
                return

            for site_index, site in enumerate(site_props, start=1):
                if self.master.cancel_requested.is_set():
                    self.log("Backup canceled by user.")
                    return

                self.global_progress["value"] = site_index
                self.global_progress_label["text"] = f"{site_index}/{total_sites} Sites Processed"
                self.log(f"Backing up site {site.url} ({site_index}/{total_sites})")

                self.backup_site(site.url)

            self.log("Backup completed successfully.")
        except Exception as e:
            self.log(f"Error: {e}")
        finally:
            self.master.is_backup_running = False

    def backup_site(self, site_url):
        try:
            site_ctx = ClientContext(site_url).with_credentials(
                UserCredential(self.master.username.get(), self.master.password.get())
            )

            web = site_ctx.web
            site_ctx.load(web)
            site_ctx.execute_query()

            site_title = web.properties.get("Title", "Untitled Site")
            local_site_folder = os.path.join(self.master.backup_folder.get(), site_title)
            os.makedirs(local_site_folder, exist_ok=True)

            lists = web.lists.filter("BaseTemplate eq 101")
            site_ctx.load(lists)
            site_ctx.execute_query()

            for sp_list in lists:
                if sp_list.properties["Title"] in SKIP_LIBRARIES:
                    self.log(f"Skipping system library: {sp_list.properties['Title']}")
                    continue

                root_folder = sp_list.root_folder
                site_ctx.load(root_folder)
                site_ctx.execute_query()

                self.log(f"Processing library: {sp_list.properties['Title']}")
                self.download_folder_recursively(site_ctx, root_folder, local_site_folder)
        except Exception as e:
            self.log(f"Error backing up site {site_url}: {e}")

    def download_folder_recursively(self, ctx, folder, local_path):
        ctx.load(folder, ["Files", "Folders", "ServerRelativeUrl"])
        ctx.execute_query()

        os.makedirs(local_path, exist_ok=True)

        for file in folder.files:
            local_file_path = os.path.join(local_path, file.name)
            try:
                self.log(f"Downloading file: {file.name}")
                response = File.open_binary(ctx, file.serverRelativeUrl)
                with open(local_file_path, "wb") as local_file:
                    local_file.write(response.content)
                self.site_progress["value"] += 1
                self.site_progress_label["text"] = f"Files Processed: {self.site_progress['value']}"
            except Exception as e:
                self.log(f"Error downloading file {file.name}: {e}")

        for subfolder in folder.folders:
            subfolder_name = subfolder.name.replace(" ", "_")
            local_subfolder_path = os.path.join(local_path, subfolder_name)
            self.download_folder_recursively(ctx, subfolder, local_subfolder_path)

if __name__ == "__main__":
    app = SharePointBackupApp()
    app.mainloop()
