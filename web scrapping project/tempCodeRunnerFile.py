df = pd.DataFrame({
    'Laptop_names': laptop_names,
    'Laptop_prices': laptop_prices,
    'Laptop_reviews': laptop_reviews
})
df.to_excel(r"C:\Users\Harshitramji11\Downloads\Learnerea\Tables.xlsx", index=False, engine='openpyxl')


driver.quit()  
# Close the browser