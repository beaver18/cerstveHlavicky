import pandas as pd
import asyncio
from playwright.async_api import async_playwright

# Načítanie e-mailov z Excelu
df = pd.read_excel("hlasovanie.xlsx")
log = []

async def hlasuj(email, heslo="MSVOLKOVCE1"):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        try:
            await page.goto("https://www.cerstvehlavicky.sk")

            # Cookies
            try:
                await page.wait_for_selector("button:has-text('Prijať vybrané')", timeout=5000)
                await page.click("button:has-text('Prijať vybrané')")
            except:
                pass

            # Prihlásenie
            await page.click("text=Prihlásiť")
            await page.fill("input[name='email']", email)
            await page.fill("input[name='password']", heslo)

            print(f"🔐 {email}: Po vyriešení CAPTCHA klikni na 'Prihlásiť sa'...")
            await page.wait_for_selector("div.vote-skolka-form", timeout=120000)

            # Vyhľadať a kliknúť na dropdown
            await page.click("div.vote-skolka-form a.chosen-single")
            await page.wait_for_selector("div.vote-skolka-form .chosen-search input", timeout=5000)
            search_input = await page.query_selector("div.vote-skolka-form .chosen-search input")
            await search_input.type("Volkovce", delay=150)
            await page.wait_for_timeout(2000)

            # Vybrať prvý výsledok
            results = await page.query_selector_all("div.vote-skolka-form .chosen-results li.active-result")
            if results:
                await results[0].click()
            else:
                raise Exception("Nenašiel sa výsledok pre 'Volkovce'")

            # Odoslať hlas
            submit_button = await page.query_selector("div.vote-skolka-form button:has-text('Odoslať hlas')")
            if submit_button:
                await submit_button.click()
                print(f"✅ {email}: Hlas odoslaný.")
            else:
                raise Exception("Tlačidlo 'Odoslať hlas' nebolo nájdené.")

            await page.wait_for_timeout(2000)
            await browser.close()
            return "hotovo"

        except Exception as e:
            print(f"❌ {email}: chyba - {e}")
            await browser.close()
            return f"chyba - {e}"

# Spustenie postupne pre všetky e-maily
async def main():
    for email in df['email']:
        if pd.isna(email):
            break
        status = await hlasuj(email)
        log.append({"email": email, "stav": status})

    # Uloženie logu
    log_df = pd.DataFrame(log)
    log_df.to_excel("log_vysledok.xlsx", index=False)
    print("📄 Log uložený do log_vysledok.xlsx")

asyncio.run(main())
