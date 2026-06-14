REQUIRED_COLUMNS = [
    'Expiration Date',
    'Insured',
    'Carrier',
    'Lines Of Business',
    'Status',
    'Premium',
    'Renewal Premium',
    'Percentage Change',
]

OUTPUT_COLUMNS = [
    'Expiration Date',
    'Insured Name',
    'Carrier',
    'Lines Of Business',
    'Status',
    'Premium',
    'Renewal Premium',
    'Percentage Change',
]

STATE_DROPDOWN = [
    'Renewal Complete',
    'Nowcerts Complete',
    'Needs Rewritten',
    'Rewritten',
    'Contact Attempted',
    'Try Bundling',
    'Already Rewritten',
    'Best Option',
    'Non Renewing',
    'Canceled',
]

CONTACTED_VIA_DROPDOWN = [
    'Left VM',
    'Sent Text',
    'Sent Email',
    'Spoken with',
]

NOTES_DROPDOWN = [
    'Yes',
    'Call Filed in AMS',
]

COMPLETED_BY_DROPDOWN = [
    'Danielle Stevens',
    'Amber Miller',
    'Teresa Morrisette',
    'Jillian Stevens',
]

STATE_FORMAT = {
    'Renewal Complete': {'bg_color': '#90EE90'},  # Light green
    'Nowcerts Complete': {'bg_color': '#36bbe9'},  # Light blue
    'Needs Rewritten': {'bg_color': '#EAE455'},  # Light yellow
    'Rewritten': {'bg_color': '#a7754d'},  # Gilmore Girl Brown
    'Contact Attempted': {'bg_color': '#9999FF'},  # Light purple
    'Try Bundling': {'bg_color': '#FFB6C1'},  # Light pink
    'Already Rewritten': {'bg_color': '#FFA500'},  # Orange
    'Best Option': {'bg_color': '#DDA0DD'},  # Plum
    'Non Renewing': {'bg_color': '#ff6666'},  # Light red
    'Canceled': {'bg_color': '#A9A9A9'},  # Dark gray
}

