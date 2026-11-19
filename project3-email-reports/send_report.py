import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import matplotlib.pyplot as plt
from dotenv import load_dotenv

# تحميل البيانات السرية من ملف .env
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

def analyze_and_chart(excel_path):
    df = pd.read_excel(excel_path)
    
    total_revenue = df["Total_Sales"].sum()
    total_orders = len(df)
    top_products = df.groupby("Product")["Total_Sales"].sum().sort_values(ascending=False).head(5)
    
    # رسم بياني لأعلى 5 منتجات
    plt.figure(figsize=(8, 4.5))
    bars = plt.bar(top_products.index, top_products.values, color="#2b5c8f", width=0.55)
    plt.title("Top 5 Best-Selling Products (Revenue)", fontsize=13, fontweight="bold", pad=12)
    plt.ylabel("Revenue (EGP)", fontsize=10)
    plt.xticks(rotation=15, ha="right")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    
    # كتابة الرقم فوق كل عمود
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + (yval * 0.02), f"{int(yval):,}", ha="center", va="bottom", fontsize=9)
        
    plt.tight_layout()
    chart_path = "sales_chart.png"
    plt.savefig(chart_path, dpi=150)
    plt.close()
    
    return total_revenue, total_orders, top_products, chart_path

def send_email_report(total_revenue, total_orders, top_products, chart_path, excel_path):
    if not SENDER_EMAIL or not APP_PASSWORD:
        print("❌ Error: SENDER_EMAIL or APP_PASSWORD missing in .env file!")
        return

    msg = EmailMessage()
    msg["Subject"] = "📊 Automated Sales Report - Performance Summary"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL

    # تجهيز جدول HTML لأعلى المنتجات
    rows_html = ""
    for product, rev in top_products.items():
        rows_html += f"<tr><td style='padding:8px; border:1px solid #ddd;'>{product}</td><td style='padding:8px; border:1px solid #ddd;'><b>{rev:,.0f} EGP</b></td></tr>"

    html_content = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
        <h2 style="color: #2b5c8f;">Automated Sales Performance Report</h2>
        <p>Hello, here is the automated summary of recent sales performance:</p>
        
        <div style="background-color: #f4f7fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <p style="margin: 5px 0; font-size: 16px;">💰 <b>Total Revenue:</b> {total_revenue:,.0f} EGP</p>
          <p style="margin: 5px 0; font-size: 16px;">📦 <b>Total Orders:</b> {total_orders} orders</p>
        </div>

        <h3>🏆 Top 5 Products by Revenue</h3>
        <table style="border-collapse: collapse; width: 60%; margin-bottom: 20px;">
          <tr style="background-color: #2b5c8f; color: white;">
            <th style="padding:8px; text-align:left;">Product</th>
            <th style="padding:8px; text-align:left;">Total Revenue</th>
          </tr>
          {rows_html}
        </table>

        <p>Please find the visual chart and raw Excel data attached.</p>
        <p style="color: #777; font-size: 12px;">Generated automatically by Python Automation Script.</p>
      </body>
    </html>
    """
    msg.set_content("Your email client does not support HTML. Please view the attached files.")
    msg.add_alternative(html_content, subtype="html")

    # إرفاق الرسم البياني وملف الإكسل
    for file_path, mime_main, mime_sub in [
        (chart_path, "image", "png"),
        (excel_path, "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ]:
        with open(file_path, "rb") as f:
            msg.add_attachment(f.read(), maintype=mime_main, subtype=mime_sub, filename=os.path.basename(file_path))

    # الإرسال عبر سيرفر Gmail
    print("⏳ Connecting to Gmail server and sending email...")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)
    print("✅ Email sent successfully to", RECEIVER_EMAIL)

if __name__ == "__main__":
    rev, orders, top5, chart = analyze_and_chart("sales_data.xlsx")
    print(f"📊 Analysis Done | Total Revenue: {rev:,.0f} EGP | Orders: {orders}")
    send_email_report(rev, orders, top5, chart, "sales_data.xlsx")