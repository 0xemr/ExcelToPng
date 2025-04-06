import pandas as pd
import matplotlib.pyplot as plt
import os
from tkinter import filedialog
import tkinter as tk

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Excel Dosyası Seçin",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path

def create_charts(data_frame, save_path):
    values = data_frame.iloc[:, 1].astype(float)
    labels = data_frame.iloc[:, 0].astype(str)
    x_pos = range(len(labels))
    
    fig = plt.figure(figsize=(15, 12))
    
    ax1 = fig.add_subplot(221)
    bars = ax1.bar(x_pos, values, color='skyblue', edgecolor='black', width=0.6)
    ax1.set_title('Çubuk Grafik', fontsize=12, pad=10)
    ax1.set_xlabel('Ürünler')
    ax1.set_ylabel('Kilogram')
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(labels, rotation=0)
    ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
    
    for bar in bars:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.1f}',
                ha='center', va='bottom')
    
    ax2 = fig.add_subplot(222)
    wedges, texts, autotexts = ax2.pie(values, labels=labels, autopct='%1.1f%%',
                                      colors=['skyblue', 'lightgreen', 'lightcoral', 'wheat'],
                                      startangle=90)
    ax2.set_title('Pasta Grafik', fontsize=12, pad=10)
    
    ax3 = fig.add_subplot(223)
    ax3.plot(labels, values, marker='o', color='green', linewidth=2, markersize=8)
    ax3.set_title('Çizgi Grafik', fontsize=12, pad=10)
    ax3.set_xlabel('Ürünler')
    ax3.set_ylabel('Kilogram')
    ax3.grid(True, linestyle='--', alpha=0.7)
    
    for i, v in enumerate(values):
        ax3.text(i, v, f'{v:.1f}', ha='center', va='bottom')
    
    ax4 = fig.add_subplot(224)
    scatter = ax4.scatter(x_pos, values, s=200, color='red', alpha=0.6)
    ax4.set_title('Nokta Grafik', fontsize=12, pad=10)
    ax4.set_xlabel('Ürünler')
    ax4.set_ylabel('Kilogram')
    ax4.set_xticks(x_pos)
    ax4.set_xticklabels(labels, rotation=0)
    ax4.grid(True, linestyle='--', alpha=0.7)
    
    for i, v in enumerate(values):
        ax4.text(i, v, f'{v:.1f}', ha='center', va='bottom')
    
    plt.tight_layout()
    
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.close()

def main():
    excel_path = select_excel_file()
    
    if not excel_path:
        print("Dosya seçilmedi!")
        return
    
    try:
        excel_dir = os.path.dirname(excel_path)
        graph_path = os.path.join(excel_dir, 'urun_grafikleri.png')
        
        df = pd.read_excel(excel_path)
        
        print("\nOkunan veriler:")
        print(df)
        
        create_charts(df, graph_path)
        
        print(f"\nGrafikler başarıyla oluşturuldu ve şu konuma kaydedildi:")
        print(graph_path)
        
    except Exception as e:
        print(f"Hata oluştu: {str(e)}")
        print("\nLütfen Excel dosyanızın şu formatta olduğundan emin olun:")
        print("1. sütun: Ürün isimleri")
        print("2. sütun: Kilogram değerleri (sayısal değerler)")

if __name__ == "__main__":
    main()
#100. satır boş kalmasın diye :D 