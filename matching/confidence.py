def confidence_score(
    customer_match: bool,
    agent_match: bool,
    delta_minutes: float,
) -> int:
    """
    امتیازدهی اعتماد (0-100):
      - تطبیق مشتری:   40 امتیاز
      - تطبیق اپراتور:  35 امتیاز
      - نزدیکی زمانی:   حداکثر 25 امتیاز

 	مشتری 	اپراتور 	زمان 	امتیاز 	Sup ≥ 75 	Rate ≥ 80
    40   	  35    25 	     100 	  ✅ 	      ✅

مشتری + اپراتور + < 5 دقیقه 	40 	35 	20 	95 	✅ 	✅
مشتری + اپراتور + < 10 دقیقه 	40 	35 	15 	90 	✅ 	✅
مشتری + اپراتور + < 15 دقیقه 	40 	35 	10 	85 	✅ 	✅
مشتری + اپراتور + < 30 دقیقه 	40 	35 	5 	80 	✅ 	✅
مشتری + اپراتور + < 45 دقیقه 	40 	35 	2 	77 	✅ 	❌
مشتری + اپراتور + > 45 دقیقه 	40 	35 	0 	75 	✅ 	❌
مشتری + بدون اپراتور + < 3 دقیقه 	40 	0 	25 	65 	❌ 	❌
مشتری + بدون اپراتور + < 5 دقیقه 	40 	0 	20 	60 	❌ 	❌
بدون مشتری + اپراتور + < 3 دقیقه 	0 	35 	25 	60 	❌ 	❌



    """
    score = 0

    if customer_match:
        score += 40

    if agent_match:
        score += 35

    abs_delta = abs(delta_minutes)
    if abs_delta < 3:
        score += 25
    elif abs_delta < 5:
        score += 20
    elif abs_delta < 10:
        score += 15
    elif abs_delta < 15:
        score += 10
    elif abs_delta < 30:
        score += 5
    elif abs_delta < 45:
        score += 2

    return score
