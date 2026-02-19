/**
 * K.303 Report Downloader
 *
 * URL format: https://www.magna.isa.gov.il/?form=%D7%A7303&q=<encoded_manager_name>
 *
 * NOTE: This site may block non-Israeli IPs. Use Israeli proxy if needed.
 */

async function pageFunction(context) {
    const { page, request, log, Apify } = context;
    const { managerName = 'unknown' } = request.userData || {};

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    log.info(`Starting K.303 download for: ${managerName}`);
    log.info(`URL: ${request.url}`);

    // Wait for Angular to fully bootstrap
    log.info('Waiting for page to fully load...');
    await sleep(5000);

    try {
        await page.waitForNetworkIdle({ idleTime: 3000, timeout: 30000 });
        log.info('Network idle reached');
    } catch (e) {
        log.warning(`Network idle timeout: ${e.message}`);
    }

    await sleep(3000);

    // Debug: Log page state
    const pageState = await page.evaluate(() => {
        const inputs = document.querySelectorAll('input');
        return {
            inputCount: inputs.length,
            reportCount: document.querySelectorAll('app-report-list-single').length,
            excelImgCount: document.querySelectorAll('img[src*="excel"]').length,
            searchValue: document.querySelector('input[placeholder*="להקליד"], input[placeholder*="חיפוש"]')?.value || ''
        };
    });
    log.info(`Page state: ${JSON.stringify(pageState)}`);

    // Check if results loaded
    let resultsFound = pageState.reportCount > 0;
    log.info(`Initial check - results found: ${resultsFound}`);

    // If no results from URL params, try manual search
    if (!resultsFound) {
        log.info('No results from URL params, attempting manual search...');

        // Find and fill search input
        const searchFilled = await page.evaluate((managerName) => {
            const input = document.querySelector('input[placeholder*="להקליד"], input[placeholder*="חיפוש"], input[type="text"]');
            if (input) {
                input.value = managerName;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.focus();
                return { success: true, placeholder: input.placeholder };
            }
            return { success: false };
        }, managerName);
        log.info(`Search input: ${JSON.stringify(searchFilled)}`);

        await sleep(2000);

        // Try autocomplete selection
        const autocomplete = await page.evaluate(() => {
            const option = document.querySelector('mat-option, [role="option"]');
            if (option) {
                option.click();
                return { selected: true, text: option.textContent?.trim()?.substring(0, 50) };
            }
            return { selected: false };
        });
        log.info(`Autocomplete: ${JSON.stringify(autocomplete)}`);

        await sleep(1000);

        // Click search button
        const searchClicked = await page.evaluate(() => {
            const btn = Array.from(document.querySelectorAll('button'))
                .find(b => b.textContent?.includes('חיפוש'));
            if (btn) {
                btn.click();
                return true;
            }
            return false;
        });
        log.info(`Search button clicked: ${searchClicked}`);

        // Wait for results
        try {
            await page.waitForNetworkIdle({ idleTime: 3000, timeout: 20000 });
        } catch (e) {
            log.warning(`Post-search network idle timeout: ${e.message}`);
        }
        await sleep(5000);
    }

    // Poll for results
    log.info('Polling for results...');
    for (let attempt = 0; attempt < 10; attempt++) {
        const count = await page.evaluate(() =>
            document.querySelectorAll('app-report-list-single').length
        );
        if (count > 0) {
            log.info(`Results found on attempt ${attempt + 1}: ${count} reports`);
            resultsFound = true;
            break;
        }
        log.info(`Attempt ${attempt + 1}: no results yet`);
        await sleep(2000);
    }

    // Save debug files
    const kvStore = await Apify.openKeyValueStore();
    const screenshot = await page.screenshot({ fullPage: true });
    await kvStore.setValue('DEBUG_screenshot', screenshot, { contentType: 'image/png' });
    const html = await page.content();
    await kvStore.setValue('DEBUG_html', html, { contentType: 'text/html' });
    log.info('Debug files saved');

    // Extract report info from listing page
    // Structure: app-report-list-single > .reportName a (link to details)
    // Dates in format: DD/MM/YYYY HH:MM
    // Report IDs in format: YYYY-NN-NNNNNN
    const reports = await page.evaluate(() => {
        const results = [];
        const reportElements = document.querySelectorAll('app-report-list-single');

        reportElements.forEach((report, idx) => {
            const nameLink = report.querySelector('.reportName a');
            const text = report.textContent || '';

            // Extract date (DD/MM/YYYY HH:MM)
            const dateMatch = text.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}:\d{2})/);
            // Extract report ID (YYYY-NN-NNNNNN)
            const idMatch = text.match(/(\d{4}-\d{2}-\d{6})/);

            if (nameLink?.href && idMatch) {
                results.push({
                    href: nameLink.href,
                    reportId: idMatch[1],
                    dateStr: dateMatch ? `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]} ${dateMatch[4]}` : '',
                    day: dateMatch ? parseInt(dateMatch[1]) : 0,
                    month: dateMatch ? parseInt(dateMatch[2]) : 0,
                    year: dateMatch ? parseInt(dateMatch[3]) : 0,
                    index: idx
                });
            }
        });

        return results;
    });

    log.info(`Found ${reports.length} reports on listing page`);

    // Check for "no results" message
    const noResults = await page.evaluate(() =>
        document.body?.innerText?.includes('לא נמצאו תוצאות')
    );
    if (noResults) {
        log.warning('Page shows "no results found" - likely geo-blocked or wrong search term');
    }

    if (reports.length === 0) {
        const bodyText = await page.evaluate(() => document.body?.innerText?.substring(0, 500));
        return {
            managerName,
            error: noResults ? 'No results found (possibly geo-blocked)' : 'No reports found',
            bodyText,
            downloads: []
        };
    }

    // Group by month and get latest from each
    const byMonth = {};
    reports.forEach(r => {
        if (r.year && r.month) {
            const monthKey = `${r.year}-${String(r.month).padStart(2, '0')}`;
            if (!byMonth[monthKey]) byMonth[monthKey] = [];
            byMonth[monthKey].push(r);
        }
    });

    const months = Object.keys(byMonth).sort().reverse();
    log.info(`Months found: ${months.join(', ')}`);

    // Download files for first 2 months
    const downloadedFiles = [];

    for (let i = 0; i < Math.min(2, months.length); i++) {
        const month = months[i];
        const monthReports = byMonth[month].sort((a, b) => b.day - a.day);
        const report = monthReports[0];
        const monthType = i === 0 ? 'current' : 'previous';

        log.info(`Processing ${monthType} month (${month}): ${report.reportId}`);

        // Navigate to report details page to find download link
        try {
            log.info(`Navigating to: ${report.href}`);
            await page.goto(report.href, { waitUntil: 'networkidle', timeout: 30000 });
            await sleep(3000);

            // Look for Excel download link on details page
            // Format: /details/downloadFile?IdFile=XXXX-xlsx-he.xlsx
            const downloadInfo = await page.evaluate(() => {
                // Try multiple selectors for the Excel download link
                const selectors = [
                    'a.download-excel',
                    'a[href*="downloadFile"][href*="xlsx"]',
                    'a[href*=".xlsx"]'
                ];

                for (const sel of selectors) {
                    const link = document.querySelector(sel);
                    if (link?.href) {
                        return { href: link.href, selector: sel };
                    }
                }

                // If no direct link, look for any downloadFile links
                const allDownloadLinks = Array.from(document.querySelectorAll('a[href*="downloadFile"]'));
                if (allDownloadLinks.length > 0) {
                    return {
                        href: allDownloadLinks[0].href,
                        selector: 'a[href*="downloadFile"]',
                        allLinks: allDownloadLinks.map(a => a.href)
                    };
                }

                return null;
            });

            if (downloadInfo?.href) {
                log.info(`Found download link: ${downloadInfo.href}`);

                // Download the file
                const response = await page.evaluate(async (url) => {
                    const resp = await fetch(url, { credentials: 'include' });
                    if (!resp.ok) return { error: `HTTP ${resp.status}` };

                    const blob = await resp.blob();
                    const reader = new FileReader();
                    return new Promise(resolve => {
                        reader.onloadend = () => resolve({
                            data: reader.result,
                            size: blob.size,
                            type: blob.type
                        });
                        reader.readAsDataURL(blob);
                    });
                }, downloadInfo.href);

                if (response.error) {
                    log.warning(`Download failed: ${response.error}`);
                    downloadedFiles.push({
                        month,
                        reportId: report.reportId,
                        error: response.error,
                        success: false
                    });
                } else if (response.data?.includes('base64,')) {
                    const filename = `k303_${monthType}_${managerName}_${month}.xlsx`;
                    const buffer = Buffer.from(response.data.split('base64,')[1], 'base64');
                    await kvStore.setValue(filename, buffer, {
                        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    });
                    log.info(`Saved ${filename} (${buffer.length} bytes)`);
                    downloadedFiles.push({
                        filename,
                        month,
                        reportId: report.reportId,
                        size: buffer.length,
                        success: true
                    });
                }
            } else {
                log.warning(`No download link found for ${report.reportId}`);

                // Save debug screenshot of details page
                const detailsScreenshot = await page.screenshot({ fullPage: true });
                await kvStore.setValue(`DEBUG_details_${report.reportId}`, detailsScreenshot, { contentType: 'image/png' });

                downloadedFiles.push({
                    month,
                    reportId: report.reportId,
                    error: 'No download link found on details page',
                    success: false
                });
            }
        } catch (e) {
            log.error(`Error processing ${report.reportId}: ${e.message}`);
            downloadedFiles.push({
                month,
                reportId: report.reportId,
                error: e.message,
                success: false
            });
        }
    }

    return {
        managerName,
        months,
        totalReports: reports.length,
        downloads: downloadedFiles
    };
}
