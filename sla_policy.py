"""Eligible SLA windows and downtime arithmetic (half-open Unix intervals)."""
from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


def validate_policy(policy):
    policy = policy or {}
    minimum = policy.get('minimum_outage_seconds', 0)
    if type(minimum) is not int or minimum < 0:
        raise ValueError('minimum_outage_seconds must be a non-negative integer')
    business = policy.get('business_hours', {}) or {}
    if business.get('enabled', False):
        days = business.get('weekdays', [0, 1, 2, 3, 4])
        if not days or any(type(day) is not int or day not in range(7) for day in days):
            raise ValueError('Business weekdays must be selected (0=Monday through 6=Sunday)')
        try:
            ZoneInfo(business.get('timezone', 'UTC'))
            start = time.fromisoformat(business.get('start', '09:00'))
            end = time.fromisoformat(business.get('end', '17:00'))
        except (ValueError, ZoneInfoNotFoundError) as exc:
            raise ValueError('Business hours require a valid timezone and HH:MM times') from exc
        if start == end:
            raise ValueError('Business start and end times must differ')
    return policy


def merge_intervals(intervals):
    merged = []
    for start, stop in sorted(intervals):
        if stop <= start:
            continue
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], stop))
        else:
            merged.append((start, stop))
    return merged


def eligible_windows(start, stop, policy):
    policy = validate_policy(policy)
    business = policy.get('business_hours', {}) or {}
    if not business.get('enabled', False):
        return [(start, stop)] if stop > start else []
    zone = ZoneInfo(business.get('timezone', 'UTC'))
    first = datetime.fromtimestamp(start, zone).date() - timedelta(days=1)
    last = datetime.fromtimestamp(stop - 1, zone).date()
    start_time = time.fromisoformat(business.get('start', '09:00'))
    end_time = time.fromisoformat(business.get('end', '17:00'))
    windows = []
    day = first
    while day <= last:
        if day.weekday() in business.get('weekdays', [0, 1, 2, 3, 4]):
            a = max(start, int(datetime.combine(day, start_time, zone).timestamp()))
            end_day = day + timedelta(days=1) if end_time < start_time else day
            b = min(stop, int(datetime.combine(end_day, end_time, zone).timestamp()))
            if b > a:
                windows.append((a, b))
        day += timedelta(days=1)
    return windows


def calculate_availability(start, stop, outages, policy=None):
    policy = validate_policy(policy)
    windows = eligible_windows(start, stop, policy)
    # Filter full outage durations before clipping to the report/business window.
    # Overlapping alerts represent one physical downtime interval.
    outages = [interval for interval in merge_intervals(outages)
               if interval[1] - interval[0] >= policy.get('minimum_outage_seconds', 0)]
    intersections = [(max(a, c), min(b, d)) for a, b in outages for c, d in windows if min(b, d) > max(a, c)]
    downtime = sum(b - a for a, b in merge_intervals(intersections))
    total = sum(b - a for a, b in windows)
    return {'availability': round((1 - downtime / total) * 100, 2) if total else None,
            'downtime_seconds':downtime, 'total_seconds':total}


def resolve_policy(defaults, overrides=None):
    defaults, overrides = defaults or {}, overrides or {}
    result = {**defaults, **overrides}
    result['business_hours'] = {**(defaults.get('business_hours', {}) or {}), **(overrides.get('business_hours', {}) or {})}
    return validate_policy(result)
