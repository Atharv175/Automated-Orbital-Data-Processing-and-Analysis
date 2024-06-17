
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import rarfile
import pandas as pd
import matplotlib.pyplot as plt

class RARDataAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RAR Data Analyzer")
        self.root.geometry("800x600")
        
        # Variables
        self.folder_path = tk.StringVar()
        self.save_dir = tk.StringVar()
        self.rarfiles = []
        
        # Main Frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=20, pady=20)
        
        # Title Label
        tk.Label(main_frame, text="RAR Data Analyzer", font=("Helvetica", 24, "bold")).grid(row=0, column=0, columnspan=4, pady=10)
        
        # Select Folder Section
        tk.Label(main_frame, text="Select Folder Containing RAR Files:", font=("Helvetica", 14)).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.folder_entry = tk.Entry(main_frame, textvariable=self.folder_path, width=50)
        self.folder_entry.grid(row=1, column=1, padx=10)
        tk.Button(main_frame, text="Browse", command=self.browse_folder).grid(row=1, column=2, padx=10)
        
        # Select Save Directory Section
        tk.Label(main_frame, text="Select Directory to Save Plots:", font=("Helvetica", 14)).grid(row=2, column=0, sticky=tk.W, pady=10)
        self.save_entry = tk.Entry(main_frame, textvariable=self.save_dir, width=50)
        self.save_entry.grid(row=2, column=1, padx=10)
        tk.Button(main_frame, text="Browse", command=self.browse_save_dir).grid(row=2, column=2, padx=10)
        
        # Process RAR Files Button
        process_button = tk.Button(main_frame, text="Process RAR Files", command=self.process_rar_files, bg="green", fg="white", font=("Helvetica", 14, "bold"))
        process_button.grid(row=3, column=1, pady=20)
        
        # Status Text
        self.status_text = tk.StringVar()
        self.status_text.set("Ready to process...")
        status_label = tk.Label(main_frame, textvariable=self.status_text, fg="blue", font=("Helvetica", 12, "italic"))
        status_label.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Progress Bar
        self.progress_bar = ttk.Progressbar(main_frame, length=500, mode='determinate')
        self.progress_bar.grid(row=5, column=0, columnspan=3, pady=10)
        
        # Suggestions Frame
        suggestions_frame = tk.Frame(main_frame, bd=2, relief=tk.GROOVE)
        suggestions_frame.grid(row=6, column=0, columnspan=3, pady=20, padx=10, sticky=tk.W+tk.E)
        
        tk.Label(suggestions_frame, text="Suggestions Based on Temperature Ranges:", font=("Helvetica", 14)).pack(pady=10)
        self.suggestions_text = scrolledtext.ScrolledText(suggestions_frame, height=10, width=70, wrap=tk.WORD, font=("Helvetica", 12))
        self.suggestions_text.pack()
        
        # Footer Label
        tk.Label(main_frame, text="© 2024 RAR Data Analyzer", fg="gray", font=("Helvetica", 10)).grid(row=7, column=0, columnspan=3, pady=10)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.rarfiles = [f for f in os.listdir(folder_selected) if f.endswith('.rar')]
            self.status_text.set(f"Found {len(self.rarfiles)} RAR files.")
    
    def browse_save_dir(self):
        save_selected = filedialog.askdirectory()
        if save_selected:
            self.save_dir.set(save_selected)
    
    def process_rar_files(self):
        folder_path = self.folder_path.get()
        save_dir = self.save_dir.get()
        
        if not folder_path or not save_dir:
            messagebox.showwarning("Warning", "Please select folder and save directory.")
            return
        
        try:
            rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"
            os.makedirs(save_dir, exist_ok=True)
            
            self.suggestions_text.delete('1.0', tk.END)  # Clear previous suggestions
            self.progress_bar['value'] = 0
            total_files = len(self.rarfiles)
            
            for i, filename in enumerate(self.rarfiles):
                rar_path = os.path.join(folder_path, filename)
                with rarfile.RarFile(rar_path) as rf:
                    times_lock = []
                    for entry in rf.infolist():
                        if entry.filename.endswith('.xlsx'):
                            if 'lock' in entry.filename:
                                times_lock.extend(self.process_lock_file(rf, entry.filename))
                    
                    for entry in rf.infolist():
                        if entry.filename.endswith('.xlsx') and 'temperature' in entry.filename:
                            filtered_temp = self.process_temperature_file(rf, entry.filename, times_lock, save_dir)
                            self.generate_suggestions(filtered_temp)
                
                # Update progress bar
                self.progress_bar['value'] = ((i + 1) / total_files) * 100
                self.root.update_idletasks()
            
            self.status_text.set("Processing complete. Plots saved.")
            messagebox.showinfo("Success", "RAR files processed successfully and plots saved.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def process_lock_file(self, rf, filename):
        times_lock = []
        with rf.open(filename) as f:
            df_lock = pd.read_excel(f)
            filtered_lock = df_lock[df_lock['lock'] == 4999.99]
            times_lock = filtered_lock['time'].tolist()
        return times_lock
    
    def process_temperature_file(self, rf, filename, times_lock, save_dir):
        with rf.open(filename) as f:
            df_temp = pd.read_excel(f)
            filtered_temp = df_temp[df_temp['time'].isin(times_lock)]
            
            if not filtered_temp.empty:
                plt.figure(figsize=(8, 4))
                plt.plot(df_temp['time'], df_temp['temperature'], marker='o', linestyle='-', label='All Data')
                plt.plot(filtered_temp['time'], filtered_temp['temperature'], marker='o', linestyle='', color='red', label='Filtered Points')
                plt.xlabel('Time')
                plt.ylabel('Temperature')
                plt.title(f'Temperature vs Time - {filename}')
                plt.legend()
                plt.grid(True)
                
                save_path = os.path.join(save_dir, f"{os.path.basename(filename).replace('.xlsx', '')}.png")
                plt.savefig(save_path)
                plt.close()
                
                return filtered_temp  # Return filtered temperature data for suggestions
            else:
                return None
    
    


    def generate_suggestions(self, df_temp):
        if df_temp is not None:
            suggestions = []
            temp_ranges = [(0, 10), (10, 20), (20, 30), (30, 40), (40, float('inf'))]  # Example temperature ranges

            for i, (low, high) in enumerate(temp_ranges):
                subset = df_temp[(df_temp['temperature'] >= low) & (df_temp['temperature'] < high)]
                count = len(subset)
                if count > 0:
                    suggestion = f"Range {i+1}: Temperature between {low} and {high} degrees, {count} data points."
                    suggestions.append(suggestion)

        # Debugging: Print suggestions to console
            print("Generated Suggestions:")
            for suggestion in suggestions:
                print(suggestion)

            self.suggestions_text.config(state=tk.NORMAL)  # Ensure the text widget is editable
            self.suggestions_text.delete('1.0', tk.END)  # Clear previous suggestions

            if suggestions:
                 for suggestion in suggestions:
                    self.suggestions_text.insert(tk.END, suggestion + "\n\n")
            else:
                self.suggestions_text.insert(tk.END, "No suggestions based on current data.\n\n")

            self.suggestions_text.config(state=tk.DISABLED)  # Make the text widget read-only
        else:
           self.suggestions_text.config(state=tk.NORMAL)  # Ensure the text widget is editable
           self.suggestions_text.delete('1.0', tk.END)  # Clear previous suggestions
           self.suggestions_text.insert(tk.END, "No temperature data available.\n\n")
           self.suggestions_text.config(state=tk.DISABLED)  # Make the text widget read-only
  



    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = RARDataAnalyzerApp(root)
    app.run()