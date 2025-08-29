#!/usr/bin/env python3

import pandas as pd
import os
import sys
from typing import Optional

class FantasyDraftManager:
    def __init__(self, spreadsheet_path: str):
        self.spreadsheet_path = spreadsheet_path
        self.sheets_data = {}
        self.load_spreadsheet()
    
    def load_spreadsheet(self):
        """Load all sheets from the Excel file"""
        try:
            print(f"Loading spreadsheet: {self.spreadsheet_path}")
            # Read all sheets
            excel_file = pd.ExcelFile(self.spreadsheet_path)
            print(f"Found sheets: {excel_file.sheet_names}")
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(self.spreadsheet_path, sheet_name=sheet_name)
                self.sheets_data[sheet_name] = df
                print(f"Loaded {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
                
        except Exception as e:
            print(f"Error loading spreadsheet: {e}")
            sys.exit(1)
    
    def display_available_players(self, sheet_name: str, max_rows: int = 20):
        """Display available players from a specific sheet"""
        if sheet_name not in self.sheets_data:
            print(f"Sheet '{sheet_name}' not found!")
            return
        
        df = self.sheets_data[sheet_name]
        print(f"\n=== {sheet_name.upper()} ===")
        print(f"Total players: {len(df)}")
        
        if len(df) == 0:
            print("No players remaining in this sheet!")
            return
        
        # Show column names
        print("\nColumns:", ", ".join([str(col) for col in df.columns.tolist()]))
        
        # Display first few rows
        print(f"\nFirst {min(max_rows, len(df))} players:")
        print(df.head(max_rows).to_string(index=False))
        
        if len(df) > max_rows:
            print(f"... and {len(df) - max_rows} more players")
    
    def search_player_by_last_name(self, last_name: str) -> list:
        """Search for players by last name across all sheets"""
        matches = []
        last_name_lower = last_name.lower().strip()
        
        for sheet_name, df in self.sheets_data.items():
            if len(df) == 0:
                continue
                
            # Look through all columns for potential player names
            for col in df.columns:
                if df[col].dtype == 'object':  # Text columns
                    # Search for last name in this column
                    mask = df[col].astype(str).str.lower().str.contains(last_name_lower, na=False, regex=False)
                    matching_rows = df[mask]
                    
                    for idx, row in matching_rows.iterrows():
                        matches.append({
                            'sheet': sheet_name,
                            'index': idx,
                            'row_data': row,
                            'matching_column': col
                        })
        
        return matches
    
    def remove_player_rows(self, matches: list, indices_to_remove: list):
        """Remove specified player rows from the spreadsheet"""
        removed_count = 0
        
        for i in indices_to_remove:
            if 0 <= i < len(matches):
                match = matches[i]
                sheet_name = match['sheet']
                row_index = match['index']
                
                # Remove the row from our data
                self.sheets_data[sheet_name] = self.sheets_data[sheet_name].drop(row_index).reset_index(drop=True)
                removed_count += 1
                print(f"Removed player from {sheet_name}")
        
        return removed_count
    
    def save_spreadsheet(self):
        """Save the updated spreadsheet"""
        try:
            # Create a backup first
            backup_path = self.spreadsheet_path.replace('.xlsx', '_backup.xlsx')
            if os.path.exists(self.spreadsheet_path):
                import shutil
                shutil.copy2(self.spreadsheet_path, backup_path)
                print(f"Backup created: {backup_path}")
            
            # Save updated data
            with pd.ExcelWriter(self.spreadsheet_path, engine='openpyxl') as writer:
                for sheet_name, df in self.sheets_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"Spreadsheet updated: {self.spreadsheet_path}")
            
        except Exception as e:
            print(f"Error saving spreadsheet: {e}")
    
    def show_menu(self):
        """Display the main menu options"""
        print("\n" + "="*50)
        print("FANTASY FOOTBALL DRAFT MANAGER")
        print("="*50)
        print("1. Show all available players")
        print("2. Show players from specific sheet")
        print("3. Remove player by last name")
        print("4. Show summary")
        print("5. Quit")
        print("-"*50)
    
    def show_summary(self):
        """Show summary of remaining players by sheet"""
        print("\n=== DRAFT SUMMARY ===")
        total_players = 0
        for sheet_name, df in self.sheets_data.items():
            count = len(df)
            total_players += count
            print(f"{sheet_name}: {count} players remaining")
        print(f"TOTAL: {total_players} players remaining")
    
    def run(self):
        """Main program loop"""
        print("Welcome to Fantasy Football Draft Manager!")
        print(f"Managing spreadsheet: {os.path.basename(self.spreadsheet_path)}")
        
        while True:
            self.show_menu()
            
            try:
                choice = input("Enter your choice (1-5): ").strip()
                
                if choice == '1':
                    # Show all players
                    for sheet_name in self.sheets_data.keys():
                        self.display_available_players(sheet_name)
                
                elif choice == '2':
                    # Show specific sheet
                    print("\nAvailable sheets:")
                    for i, sheet_name in enumerate(self.sheets_data.keys(), 1):
                        print(f"{i}. {sheet_name}")
                    
                    try:
                        sheet_choice = int(input("Enter sheet number: ")) - 1
                        sheet_names = list(self.sheets_data.keys())
                        if 0 <= sheet_choice < len(sheet_names):
                            self.display_available_players(sheet_names[sheet_choice])
                        else:
                            print("Invalid sheet number!")
                    except ValueError:
                        print("Please enter a valid number!")
                
                elif choice == '3':
                    # Remove player
                    last_name = input("Enter player's last name to remove: ").strip()
                    if not last_name:
                        print("Please enter a last name!")
                        continue
                    
                    matches = self.search_player_by_last_name(last_name)
                    
                    if not matches:
                        print(f"No players found with last name '{last_name}'")
                        continue
                    
                    print(f"\nFound {len(matches)} potential matches:")
                    for i, match in enumerate(matches):
                        row_data = match['row_data']
                        print(f"{i+1}. Sheet: {match['sheet']}")
                        print(f"   Data: {row_data.to_dict()}")
                        print()
                    
                    # Ask which ones to remove
                    remove_input = input("Enter numbers to remove (e.g., 1,3 or 'all' or 'none'): ").strip().lower()
                    
                    if remove_input == 'none':
                        continue
                    elif remove_input == 'all':
                        indices_to_remove = list(range(len(matches)))
                    else:
                        try:
                            indices_to_remove = [int(x.strip()) - 1 for x in remove_input.split(',') if x.strip()]
                        except ValueError:
                            print("Invalid input format!")
                            continue
                    
                    if indices_to_remove:
                        removed_count = self.remove_player_rows(matches, indices_to_remove)
                        if removed_count > 0:
                            print(f"Removed {removed_count} player(s)")
                            self.save_spreadsheet()
                        else:
                            print("No players were removed")
                
                elif choice == '4':
                    # Show summary
                    self.show_summary()
                
                elif choice == '5':
                    # Quit
                    print("Thanks for using Fantasy Football Draft Manager!")
                    break
                
                else:
                    print("Invalid choice! Please enter 1-5.")
            
            except KeyboardInterrupt:
                print("\n\nGoodbye!")
                break
            except Exception as e:
                print(f"An error occurred: {e}")
                print("Please try again.")

def main():
    spreadsheet_path = "/Users/jjagt/hacker/ff-moneyball/FF - main.xlsx"
    
    if not os.path.exists(spreadsheet_path):
        print(f"Error: Spreadsheet not found at {spreadsheet_path}")
        return
    
    manager = FantasyDraftManager(spreadsheet_path)
    manager.run()

if __name__ == "__main__":
    main()
