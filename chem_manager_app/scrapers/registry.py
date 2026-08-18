"""Single source of truth for vendor normalization and scraper routing."""

VENDORS = {
    "Aldrich": ("aldrich", "sigma", "merck", "sial", "millipore", "머크", "알드리치", "시그마"),
    "ThermoFisher": ("thermo", "alfa", "fisher", "invitrogen", "acros", "써모", "피셔"),
    "TCI": ("tci", "tokyo kasei", "tokyo", "티씨아이", "도쿄카세이"),
    "Abcam": ("abcam", "앱캠"),
}


def normalize_manufacturer(value):
    if not value:
        return ""
    text = str(value).strip()
    lowered = text.casefold()
    for canonical, aliases in VENDORS.items():
        if any(alias in lowered for alias in aliases):
            return canonical
    return text.title() if len(text) > 1 else text.upper()


def product_key(manufacturer, catalog_no):
    manufacturer = normalize_manufacturer(manufacturer).casefold()
    catalog = str(catalog_no or "").strip()
    if catalog.endswith(".0"):
        catalog = catalog[:-2]
    return manufacturer, catalog.casefold()


def scraper_class(manufacturer):
    canonical = normalize_manufacturer(manufacturer)
    if canonical == "Aldrich":
        from scrapers.aldrich import AldrichScraper
        return AldrichScraper
    if canonical == "ThermoFisher":
        from scrapers.thermofisher import ThermofisherScraper
        return ThermofisherScraper
    if canonical == "TCI":
        from scrapers.tci import TciScraper
        return TciScraper
    if canonical == "Abcam":
        from scrapers.abcam import AbcamScraper
        return AbcamScraper
    return None


def create_scraper(manufacturer, **kwargs):
    cls = scraper_class(manufacturer)
    return cls(**kwargs) if cls else None
