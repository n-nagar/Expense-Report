# test_scraper.py
# This is a standalone script to test the web scraping functionality.
# Run it from your terminal: python test_scraper.py

import utils
import calendar

def test_mumbai_rates():
    """
    Tests the web scraper by fetching and printing the per diem rates
    for Mumbai in May 2025.
    """
    print("--- Testing Per Diem Scraper ---")
    year = 2025
    month = 5 # May
    
    # Get all rates for India in May 2025
    all_rates = utils.get_per_diem_rates_with_selenium(year, month)
    
    if all_rates:
        # The website uses "mumbai (bombay)" as the key
        for city in all_rates.items():
            print(f"\n--- Scraped Rates for {city[0]}-")
            print(f"Lodging: ${city[1]['lodging']}")
            print(f"Total M&IE: ${city[1]['total_mie']}")
            
            # Now, let's get the breakdown
            mie_breakdown = utils.get_mie_breakdown()
            if mie_breakdown:
                print("\nM&IE Breakdown:")
                print(f"  - Breakfast: ${mie_breakdown['breakfast']}")
                print(f"  - Lunch: ${mie_breakdown['lunch']}")
                print(f"  - Dinner: ${mie_breakdown['dinner']}")
                print(f"  - Incidentals: ${mie_breakdown['incidentals']}")
            else:
                print(f"\nCould not find M&IE breakdown for a total of ${city[1]['total_mie']}.")
    else:
        print("\nFailed to scrape any per diem rates.")

if __name__ == "__main__":
    test_mumbai_rates()
