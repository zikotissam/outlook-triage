CANONICAL = [
    "important",
    "ads",
    "security",
    "urgent",
    "finance",
    "travel",
    "personal",
    "updates",
    "other",
]


def outlook_label_category(inference_classification: str | None, categories: list[str] | None = None) -> str | None:
    """Outlook signals that hint at bulk/ads vs updates.

    `inferenceClassification` is Outlook's "Focused Inbox" signal. `other`
    strongly suggests bulk/ads mail; it can't distinguish ads from updates on
    its own, so callers pair it with bulk headers to pick the category.
    """
    categories = categories or []
    for cat in categories:
        lowered = cat.lower()
        if "promotion" in lowered or "advert" in lowered or "marketing" in lowered or "newsletter" in lowered:
            return "ads"
    if inference_classification == "other":
        return "other"
    return None
