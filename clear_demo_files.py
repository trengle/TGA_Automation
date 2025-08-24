import os
import glob

def main():
    csv_files = glob.glob("*.csv")
    xl_files = glob.glob("*.xlsx")
    demo_files = csv_files + xl_files   
    if not csv_files or not xl_files:
        print("No csv or xl files found.")
    else:
        confirm = input("Are you sure you want to delete files? (yes/no): ")
        if confirm.lower() == 'yes':
            for file in demo_files:
                # print(file)           # Dry run
                os.unlink(file)
                print(f"{file} deleted") # Perform deletion
            pass
        else:
            print("Deletion aborted.")

if __name__ == "__main__":
    main()

    
