import os

# files = ["A Average Citizenship Honor Roll","Citizenship Honor Roll","Principal Honor Roll","Regular Honor Roll","Superior Honor Roll", "totals"]

# filename = '5052Q3(1)'
# save_directory = "C:\\Users\\ariel\\Desktop\\Python Projects\\Excel\\HonourRoll2\\"

# def file_clean_up():
#     for item in files:
#         item = item.replace(" ","").lower() + "_" + filename
#         path = os.path.join(save_directory,item)
#         print(path)
#         print(os.path.exists(path))
#         if os.path.exists(path):
#             os.remove(f"{item}.txt")
#     print("Cleaned up!")


# print(os.path.exists("C:\\Users\\ariel\\Desktop\\Python Projects\\Excel\\HonourRoll2\\superiorhonorroll_5052Q3(1).txt"))

with open("C:\\Users\\ariel\\Desktop\\Python Projects\\Excel\\HonourRoll2\\principalhonorroll_5052Q3(1).txt") as f:
    for line in f:
        updated_list = ' '.join(line.split("  ")).split()
        print(updated_list)
            