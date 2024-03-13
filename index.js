const puppeteer = require('puppeteer');
const path = require('path');
const FolderName = "Bundle G";
main()
async function main() {

    // Launch a headless browser
    const browser = await puppeteer.launch({
        headless: false, // Set to true for headless mode
        userDataDir: './outlook-session', // Directory to store session data
        default_directory: path.join(process.cwd(), '/'),
    });

    // Open a new page
    const page = await browser.newPage();
    await page.setViewport({ width: 1600, height: 900 });
    // Navigate to Outlook login page
    await page.goto('https://outlook.office365.com/mail/');

    // Wait for the login page to load
    await page.waitForSelector('#MailList');

    setInterval(() => { CheckMain(page); }, 10000)



}

async function CheckMain(page) {
    const elements = await page.$$('div[aria-label*="Unread"]');

    for (const element of elements) {
        // Extract the aria-label attribute
        const ariaLabel = await element.evaluate(element => element.getAttribute('aria-label'));

        // Check if the word 'unread' exists in the aria-label attribute
        if (ariaLabel.toLowerCase().includes('unread')) {
            console.log("Found element with 'unread' in the aria-label attribute:", ariaLabel);

            // Click on the element
            await element.click();
            await element.press('Enter');
            console.log("Clicked on the element.");

            // Wait for the listbox parent element to appear
            await page.waitForSelector('[role="listbox"]');
            await page.waitForSelector('[data-icon-name="pdf20_svg"]');

            await page.click('[data-icon-name="pdf20_svg"]');

            // Wait for the button with aria-label "Download" to appear
            await page.waitForSelector('button[aria-label="Download"]');

            const client = await page.target().createCDPSession();
            await client.send("Page.setDownloadBehavior", {
                behavior: "allow",
                downloadPath: path.resolve(__dirname, FolderName)
            });

            // Click on the button
            await page.click('button[aria-label="Download"]');
            console.log("Clicked on the button with aria-label='Download'.");

            await delay(1000);

            // Click on the button
            await page.click('button[aria-label="Close"]');
            console.log("Clicked on the button with aria-label='Close'.");

            await delay(1000);

            await page.waitForSelector('[data-icon-name="MailReadRegular"]');

            await page.click('[data-icon-name="MailReadRegular"]');

        }
    }
}


function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
