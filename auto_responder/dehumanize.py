from datetime import datetime, timedelta,date
import re


def naturaltime(text, now=None):
    """Convert a naturaltime string to a datetime object."""
    text = text.lower().strip()
    if not now:
        now = datetime.now()

    if text == 'now':
        return now
    if text == "today":
        return date.today()
    if text == 'yesterday':
        return now - timedelta(days=1)
    if text == 'tomorrow':
        return now + timedelta(days=1)

    if "ago" or "last" in text:
        multiplier = -1
    elif "from now" in text:
        multiplier = 1
    else:
        raise ValueError("%s is not a valid naturaltime" % text)

    text = text.replace('an ', '1 ')
    text = text.replace('a ', '1 ')

    if "last" in text:
        last_reg = re.compile(re.escape('last '), re.IGNORECASE)
        text = last_reg.sub('1 ', text)

    years = get_first(r'(\d*) year', text)
    months = get_first(r'(\d*) month', text)
    weeks = get_first(r'(\d*) week', text)
    days = get_first(r'(\d*) day', text)
    days = days + weeks * 7 + months * 30 + years * 365

    hours = get_first(r'(\d*) hour', text)
    minutes = get_first(r'(\d*) minute', text)
    seconds = get_first(r'(\d*) second', text)
    delta = timedelta(
        days=days,
        hours=hours,
        minutes=minutes,
        seconds=seconds)
    delta *= multiplier
    return now + delta


def get_first(pattern, text):
    """Return either a matched number or 0."""
    matches = re.findall(pattern, text)
    if matches:
        return int(matches[0])
    else:
        return 0

if __name__ == "__main__":
    def test_parse_time_past():
        now = datetime.now()
        tests = [
            ('57 seconds ago', -timedelta(seconds=57)),
            ('4 minutes ago', -timedelta(minutes=4)),
            ('4 hours ago', -timedelta(hours=4)),
            ('an hour ago', -timedelta(hours=1)),
            ('a minute ago', -timedelta(minutes=1)),
            ('now', -timedelta(hours=0)),
            ('1 day ago', -timedelta(days=1)),
            ('2 days ago', -timedelta(days=2)),
            ('1 day ago', -timedelta(days=1)),
            ('1 day, 1 hour ago', -timedelta(days=1, hours=1)),
            ('1 day, 2 hours ago', -timedelta(days=1, hours=2)),
            ('2 days, 1 hour ago', -timedelta(days=2, hours=1)),
            ('2 days, 2 hours ago', -timedelta(days=2, hours=2)),
            ]
        for input_time, delta in tests:
            parsed = naturaltime(input_time, now)
            expected = (now + delta)
            assert(parsed == expected)
    test_parse_time_past()
    print(naturaltime("today"))