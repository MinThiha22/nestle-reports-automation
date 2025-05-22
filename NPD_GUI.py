import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
from tkcalendar import DateEntry
from tkinter import ttk
import threading
import time
from NPD import automate_excel_process, prevent_sleep, allow_sleep
import logging
import os
import sys

class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.configure(state='normal')
        self.widget.insert('end', s)
        self.widget.see('end')
        self.widget.configure(state='disabled')

    def flush(self):
        pass

class NPD_GUI:
    def __init__(self, root):
        self.stopped = False
        self.root = root
        root.title("NPD Automation")
        # Center the window on the screen
        width = 600
        height = 600
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        
        root.resizable(False, False)
        # Set window icon (with error handling)
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
            else:
                icon_path = 'icon.ico'
            
            # Verify the file exists before trying to load it
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
            else:
                print(f"Icon file not found at: {icon_path}")  # Debug message
        except Exception as e:
            print(f"Could not load window icon: {e}")  # Debug message
          # Set up logging           
        self.create_file_browse_frame()
        self.create_schdule_frame()
        tk.Label(root, text="---OR---").pack(pady=5)
        self.start_now_frame()
        self.output_widgets()
        
        self.bind_close_event()

    def create_file_browse_frame(self):
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=20)
        
        self.file_label = tk.Label(file_frame, text="Select Pivot File: ")
        self.file_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.file_path_var = tk.StringVar()
        self.file_path_entry = tk.Entry(file_frame, textvariable=self.file_path_var, width=55)
        self.file_path_entry.pack(side=tk.LEFT, padx=(0, 5))

        self.browse_button = tk.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(side=tk.RIGHT, padx=(5, 0))

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set("File selected")
            self.start_button.config(state=tk.NORMAL)
            self.save_schedule_button.config(state=tk.NORMAL)
    
    def create_schdule_frame(self):
        schedule_frame = tk.Frame(self.root)
        schedule_frame.pack(pady=10)
        
        self.schedule_var = tk.StringVar()
        self.schedule_label = tk.Label(schedule_frame, text="Schedule Refresh: ")
        self.schedule_label.pack(side=tk.LEFT)

        self.date_entry = DateEntry(
            schedule_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='y-mm-dd')
        self.date_entry.pack(side=tk.LEFT, padx=(5, 5))

        # Hour dropdown (00-23)
        # Set hour_var to current hour as default
        current_hour = time.strftime("%H")
        self.hour_var = tk.StringVar(value=current_hour)
        self.hour_combo = ttk.Combobox(schedule_frame, textvariable=self.hour_var, width=3, values=[f"{i:02d}" for i in range(24)], state="readonly")
        self.hour_combo.pack(side=tk.LEFT, padx=(5, 0))

        tk.Label(schedule_frame, text=":").pack(side=tk.LEFT)

        # Minute dropdown (00-59)
        current_minute = time.strftime("%M")
        self.minute_var = tk.StringVar(value=current_minute)
        self.minute_combo = ttk.Combobox(schedule_frame, textvariable=self.minute_var, width=3, values=[f"{i:02d}" for i in range(60)], state="readonly")
        self.minute_combo.pack(side=tk.LEFT, padx=(0, 0))

        self.time_label = tk.Label(schedule_frame, text="(HH:MM, 24h)")
        self.time_label.pack(side=tk.LEFT, padx=(5, 0))

        # Save button to store date and time
        self.save_schedule_button = tk.Button(schedule_frame, text="Save Schedule", command=self.save_schedule)
        self.save_schedule_button.pack(side=tk.LEFT, padx=(10, 0))
        self.save_schedule_button.config(state=tk.DISABLED)

        # Label to display saved schedule time
        self.saved_schedule_label = tk.Label(self.root, text="No schedule set.", fg="green")
        self.saved_schedule_label.pack(pady=(0, 5))
        
        self.cancel_schedule_button = tk.Button(self.root, text="Cancel Schedule", command=self.cancel_schedule)
        self.cancel_schedule_button.pack(pady=(0, 5))
        self.cancel_schedule_button.config(state=tk.DISABLED)

    def save_schedule(self):
        if not self.file_path_var.get():
            tk.messagebox.showerror("Error", "Please select a file first.")
            return
        selected_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        selected_hour = self.hour_var.get()
        selected_minute = self.minute_var.get()
        schedule_str = f"{selected_date} {selected_hour}:{selected_minute}"
        

        if tk.messagebox.askyesno(
            "Confirm Schedule",
            f"Do you want to save this schedule?\n\n{schedule_str}"
        ):
            
            is_valid = self.schedule_automation(selected_date, selected_hour, selected_minute)
            if is_valid:
                self.cancel_schedule_button.config(state=tk.NORMAL)
                self.start_button.config(state=tk.DISABLED)
                tk.messagebox.showinfo("Info", "Schedule saved successfully.")
                self.status_var.set("Waiting for scheduled time...")
                self.saved_schedule_label.config(
                text=f"Schedule saved. Schedule refresh will start at {schedule_str}")
            
                if tk.messagebox.askyesno(
                    "Minimise to Background",
                    "Do you want to minimise the app and run automation in the background until the scheduled time?"
                ):
                    self.root.iconify()
        else:
            self.saved_schedule_label.config(text="Schedule not saved.")
            self.saved_schedule_label.config(fg="red")
    
    def schedule_automation(self, date, hour, minute):
        # Convert the date and time to a timestamp
        schedule_time = f"{date} {hour}:{minute}"
        schedule_timestamp = time.mktime(time.strptime(schedule_time, "%Y-%m-%d %H:%M"))

        # Calculate the delay until the scheduled time
        current_time = time.time()
        delay = schedule_timestamp - current_time

        if delay > 0:
            thread = threading.Timer(delay, self.start_automation)
            thread.name = "ScheduledAutomation"
            thread.start()
            self.saved_schedule_label.config(text=f"Scheduled automation for {schedule_time}")
            self.saved_schedule_label.config(fg="green")
            return True
        else:
            tk.messagebox.showerror("Error", "Scheduled time is in the past.")
            return False
        
    def cancel_schedule(self):
        # Cancel any existing scheduled automation
        for thread in threading.enumerate():
            if thread.name == "ScheduledAutomation":
                thread.cancel()
                self.saved_schedule_label.config(text="Scheduled automation cancelled.")
                self.saved_schedule_label.config(fg="red")
                self.cancel_schedule_button.config(state=tk.DISABLED)
                tk.messagebox.showinfo("Info", "Scheduled cancelled.")
                self.status_var.set("Scheduled cancelled")
                return
        tk.messagebox.showinfo("Info", "No scheduled automation to cancel.")
    
    def start_now_frame(self):
        start_now_frame = tk.Frame(self.root)
        start_now_frame.pack(pady=10)
        
        self.start_button = tk.Button(start_now_frame, text="Start Automation Now", command=self.start_automation)
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        self.start_button.config(state=tk.DISABLED)
        
        self.stop_button = tk.Button(start_now_frame, text="Stop Automation", command=self.stop_automation)
        self.stop_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.stop_button.config(state=tk.DISABLED)
    
    def start_automation(self):
        NPD_GUI.stop_requested = False
        self.stopped = False
        if not self.file_path_var.get():
            tk.messagebox.showerror("Error", "Please select a file.")
            return
        self.error_filename = "error.log"
        self.status_var.set("Running...")
        self.timer_var.set("⏱️ Timer: 00:00")
        self.file_path_entry.config(state=tk.DISABLED)
        self.browse_button.config(state=tk.DISABLED)
        self.save_schedule_button.config(state=tk.DISABLED)
        self.cancel_schedule_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        self.date_entry.config(state=tk.DISABLED)
        self.hour_combo.config(state=tk.DISABLED)
        self.minute_combo.config(state=tk.DISABLED)
        
        self.start_time = time.time()
        self.timer_running = True
        self.update_timer()
        
        threading.Thread(target=self.run_automation, daemon=True).start()
        
    def run_automation(self):
        try:
            prevent_sleep()
            automate_excel_process(self.file_path_var.get())
            if self.stopped:
                self.status_var.set("🎯 Ready")
            else:
                self.status_var.set("✅ Success!")
        except Exception as e:
            self.status_var.set(f"❌ Error: {e}")
        finally:
            allow_sleep()
            logging.shutdown()
            print(f"Log file: {self.error_filename}")  # <- fixed
            error_filename = self.error_filename       # <- fixed
            if (os.path.exists(error_filename) and os.path.getsize(error_filename) == 0) or self.stopped:
                os.remove(error_filename)
            else:
                print(f"📝 Errors logged in: {error_filename}")
            
            self.file_path_entry.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.NORMAL)
            self.save_schedule_button.config(state=tk.NORMAL)
            
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.date_entry.config(state=tk.NORMAL)
            self.hour_combo.config(state=tk.NORMAL)
            self.minute_combo.config(state=tk.NORMAL)
            
            self.console.configure(state='normal')
            self.console.delete(1.0, tk.END)
            self.console.configure(state='disabled')
            self.timer_running = False
        
    def stop_automation(self):
        if self.start_time is None:
            tk.messagebox.showinfo("Info", "No automation process to stop.")
            return
        if tk.messagebox.askyesno("Confirm Stop", "Are you sure you want to stop the automation?"):
            NPD_GUI.stop_requested = True
            self.stopped = True
            print("Automation stop by user...")
            self.status_var.set("Stopping...")
            self.file_path_entry.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.NORMAL)
            self.save_schedule_button.config(state=tk.NORMAL)
            
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.date_entry.config(state=tk.NORMAL)
            self.hour_combo.config(state=tk.NORMAL)
            self.minute_combo.config(state=tk.NORMAL)
            
            self.timer_running = False
            self.timer_var.set("⏱️ Timer: 00:00")
            self.start_time = None
            
        
    def output_widgets(self):
        self.status_var = tk.StringVar()
        self.status_var.set("🎯 Ready")
        self.status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI Emoji", 14, "bold"),
            relief="solid",
            fg="green",
            borderwidth=1,
            padx=10,
            pady=5
        )
        self.status_label.pack(pady=5)
        
        self.timer_var = tk.StringVar()
        self.timer_var.set("⏱️ Timer: 00:00")
        self.start_time = None
        self.timer_running = False
        
        self.timer_label = tk.Label(root, textvariable=self.timer_var, font=("Arial", 12))
        self.timer_label.pack(pady=5)

        self.console = scrolledtext.ScrolledText(root, height=15, width=70, state='disabled', font=("Consolas", 9), bg="black", fg="white")
        self.console.tag_configure("error", foreground="red")
        self.console.pack(pady=5)
        
        # Redirect stdout to the console
        sys.stdout = TextRedirector(self.console)
        sys.stderr = TextRedirector(self.console)

    def update_timer(self):
        if self.timer_running:
            elapsed = int(time.time() - self.start_time)
            mins, secs = divmod(elapsed, 60)
            self.timer_var.set(f"⏱️ Timer: {mins:02d}:{secs:02d}")
            self.root.after(1000, self.update_timer)

    def on_closing(self):
        running = self.timer_running
        scheduled = any(thread.name == "ScheduledAutomation" for thread in threading.enumerate())
        if running or scheduled:
            msg = "Automation is currently running." if running else ""
            if scheduled:
                msg += "\nA scheduled automation is pending."
            msg += "\nAre you sure you want to exit?"
            if not tk.messagebox.askyesno("Confirm Exit", msg):
                return
        self.root.destroy()

    # Bind the close event
    def bind_close_event(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
       
if __name__ == "__main__":
    root = tk.Tk()
    app = NPD_GUI(root)
    root.mainloop()