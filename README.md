# Vehicle Emissions Calculator (2024)

A **Google Script** I developed during an internship to estimate the transportation emissions of participants travelling to and from an eco-focused chapter of a global networking organisation.

# Disclaimer

This app does not work on local development, as this is a crude lift-and-shift code transfer. I created this app long before I knew about git (which would've saved me tons of time), and I have no plans of making any future adjustments or edits. I will be focused on creating other, more improved emission calculators in the future.

## API Tools

This app uses the **Google Maps API** to measure travel distances and times for each participant.

Previous demos estimated vehicle emissions using APIs including _Carbon Sutra_ and _Carbon Interface_, however many car makes weren't listed. 

Instead, I use emission factors based on fuel types drawn from reputable sources including:

- [Razif 2016](https://www.researchgate.net/publication/301678822_Prediction_of_CO_CO2_CH4_and_N2O_vehicle_emissions_from_environmental_impact_assessment_EIA_at_toll_road_of_Krian-Legundi-Bunder_in_East_Java_of_Indonesia) (for gas and diesel cars)
- [Ritchie 2023](https://ourworldindata.org/travel-carbon-footprint) (for petrol cars)
- [IEA](https://www.iea.org/countries/united-arab-emirates/electricity) (for electric cars)
- [Dubai Statistics Centre 2021](https://www.dsc.gov.ae/en-us/Themes/Pages/Transport.aspx?Theme=31) (for metro data)
## Features

- Argues with survey results on Google Sheets (from Typeform; reccomended for conditional questioning)
- Automatically processes new results every month
- Emails each participant with their own results
- Archives the results
- Estimates the equivalent number of mangroves needed to plant to outweigh the emissions.
