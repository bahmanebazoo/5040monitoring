def confidence_score(
    customer_match: bool,
    agent_match: bool,
    delta_minutes: float
) -> int:
    score = 0

    if customer_match:
        score += 40

    if agent_match:
        score += 30

    if abs(delta_minutes) < 5:
        score += 20
    elif abs(delta_minutes) < 15:
        score += 10
    elif abs(delta_minutes) < 40:
        score += 5

    return score
