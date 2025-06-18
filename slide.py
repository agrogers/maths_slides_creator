import random
from math import gcd
from fractions import Fraction
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import re
from fractions import Fraction
from sympy import isprime

# Simple color name to RGB map
COLOR_MAP = {
    "black": RGBColor(0, 0, 0),
    "blue": RGBColor(0, 0, 255),
    "red": RGBColor(255, 0, 0),
    "green": RGBColor(0, 128, 0),
    "dark_gray": RGBColor(100, 100, 100),
}
fractions_unicode = {
    "1/2": "½",
    "1/3": "⅓", "2/3": "⅔",
    "1/4": "¼", "3/4": "¾",
    "1/5": "⅕", "2/5": "⅖", "3/5": "⅗", "4/5": "⅘",
    "1/6": "⅙", "5/6": "⅚",
    "1/7": "⅐", "2/7": "²⁄₇", "3/7": "³⁄₇", "4/7": "⁴⁄₇", "5/7": "⁵⁄₇", "6/7": "⁶⁄₇",
    "1/8": "⅛", "3/8": "⅜", "5/8": "⅝", "7/8": "⅞"
}
question_type = {
    1: {
        "add": [
            {"qty":30, "min": 0, "max": 7, "tiers": range(1,2)},
            {"qty":30, "min": 5, "max": 20, "tiers": range(2,3)}
        ],
        "subtract": [{"qty":20, "min": 0, "max": 20, "tiers": range(3,5)}],
        "multiply": [{"qty":20, "min": 1, "max": 10, "tiers": range(3,6)}],
        "place_value": [{"qty":10, "min": 1, "max": 99, "tiers": range(3,5), "fontsize": 80}],
        "place_value_reverse": [{"qty":10, "min": 10, "max": 99, "tiers": range(3,6), "fontsize": 80}],
    },
    2: {
        "add": [
            {"qty":30, "min": 0, "max": 10, "tiers": range(1,2)},
            {"qty":15, "min": 5, "max": 40, "tiers": range(2,3)}
        ],
        "add3": [
            {"qty":15, "min": 1, "max": 10, "tiers": range(4,6)}],
        "subtract": [
            {"qty":15, "min": 0, "max": 20, "tiers": range(2,3)},
            {"qty":20, "min": 0, "max": 20, "tiers": range(3,5)}
        ],
        "multiply": [{"qty":20, "min": 1, "max": 10, "tiers": range(3,6)}],
        "place_value": [{"qty":10, "min": 1, "max": 999, "tiers": range(3,5), "fontsize": 80}],
        "place_value_reverse": [{"qty":10, "min": 10, "max": 9999, "tiers": range(3,6), "fontsize": 80}],
    },
    3: {
        "add": [{"qty":20, "min": 10, "max": 100, "tiers": range(1,5)}],
        "subtract": [{"qty":20, "min": 10, "max": 100, "tiers": range(2,5)}],
        "multiply": [{"qty":30, "min": 1, "max": 50, "tiers": range(3,5)}],
        "divide": [{"qty":15, "min": 1, "max": 50, "tiers": range(3,5)}],
    },
    4: {
        "add": [
            {"qty":15, "min": 10, "max": 30, "tiers": range(1,2)},
            {"qty":10, "min": 100, "max": 1000, "tiers": range(3,4)},
            {"qty":10, "min": 1000, "max": 5000, "tiers": range(4,5)},
            ],
        "add3": [
            {"qty":15, "min": 5, "max": 20, "tiers": range(2,3)}],
        "add_fraction_same_denominator": [
            {"qty":10, "min": 20, "max": 99, "tiers": range(3,5)}],
        "add_fraction_different_denominator": [
            {"qty":10, "min": 20, "max": 99, "tiers": range(4,6)}],
        "subtract": [
            {"qty":15, "min": 10, "max": 50, "tiers": range(1,2)},
            {"qty":15, "min": 100, "max": 1000, "tiers": range(3,5)},
            ],
        "multiply": [
            {"qty":15, "min": 3, "max": 7, "tiers": range(2,3)},
            {"qty":15, "min": 5, "max": 10, "tiers": range(3,5)},
            ],
        "divide": [
            {"qty":20, "min": 2, "max": 100, "tiers": range(3,6)}],
        "place_value_reverse": [
            {"qty":15, "min": 10, "max": 9999, "tiers": range(3,5), "fontsize": 70}],
    },
    5: {
        "add": [
            {"qty":15, "min": 10, "max": 30, "tiers": range(1,2)},
            {"qty":15, "min": 20, "max": 100, "tiers": range(2,3)},
            {"qty":10, "min": 100, "max": 1000, "tiers": range(3,4)},
            {"qty":10, "min": 1000, "max": 5000, "tiers": range(4,5)},
            ],
        "add3": [{"qty":10, "min": 20, "max": 99, "tiers": range(3,5)}],
        "add_dec": [{"qty":10, "min": 0, "max": 20, "tiers": range(3,5)}],
        "subtract": [
            {"qty":15, "min": 20, "max": 100, "tiers": range(2,3)},
            {"qty":15, "min": 100, "max": 1000, "tiers": range(3,5)},
            ],
        "multiply": [
            {"qty":15, "min": 3, "max": 7, "tiers": range(2,3)},
            {"qty":15, "min": 5, "max": 10, "tiers": range(3,5)},
            ],
        "divide": [{"qty":15, "min": 2, "max": 100, "tiers": range(3,5)}],
    },

}

def add_textbox(slide, text, left, top, width, height, default_font_size=44, 
                default_color = COLOR_MAP['black'], bold=True, 
                align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE):
    # txBox = slide.shapes.add_textbox(left, top, width, height)
    # tf = txBox.text_frame
    # p = tf.paragraphs[0]
    # run = p.add_run()
    # run.text = text
    # font = run.font
    # font.size = Pt(font_size)
    # font.bold = bold
    # p.alignment = align
    # if valign:
    #     tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    # return txBox

    def parse_size(size_str, default_size_pt):
        if not size_str:
            return default_size_pt
        if size_str.endswith('%'):
            percent = float(size_str[:-1])
            return default_size_pt * percent / 100
        else:
            return float(size_str)

    def add_run(paragraph, text, size, color):
        run = paragraph.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.color.rgb = color
        return run

    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True  # This is default, but explicitly set it
    text_frame.clear()
    p = text_frame.paragraphs[0]
    p.line_spacing = 0.8  # proportion, not points

    pattern = re.compile(
        r'<font(?:\s+size=([\d]+%?))?(?:\s+color=(\w+))?\s*>(.*?)</font>',
        re.DOTALL
    )

    
    matches = list(pattern.finditer(text))
    last_end = 0

    for match in matches:
        # Default-formatted text before the match
        if match.start() > last_end:
            add_run(p, text[last_end:match.start()], default_font_size, default_color)

        size_str = match.group(1)
        color_str = match.group(2)
        content = match.group(3)

        size = parse_size(size_str, default_font_size)
        color = COLOR_MAP.get(color_str.lower(), default_color) if color_str else default_color
        content = match.group(3)

        add_run(p, content, size, color)
        last_end = match.end()

    # Remaining plain text after last match
    if last_end < len(text):
        add_run(p, text[last_end:], default_font_size, default_color)

    p.alignment = align
    text_frame.vertical_anchor = valign

    return textbox

def generate_number(min_val, max_val, tier, tier_range):
    """Generate a number scaled by tier difficulty, including min_val and max_val."""
    # Total number of tiers
    num_tiers = len(tier_range)

    # Size of each tier range
    range_size = (max_val - min_val + 1) / num_tiers  # +1 to include max_val

    # Determine this tier's range
    tier_index = tier - min(tier_range)
    tier_min = int(min_val + tier_index * range_size)
    tier_max = int(min_val + (tier_index + 1) * range_size - 1)

    # Clamp to global min and max to avoid overshooting
    tier_min = max(min_val, tier_min)
    tier_max = min(max_val, tier_max)

    return random.randint(tier_min, tier_max)

def simplify_fraction(num, denom):
    """Simplify a fraction and return as a string."""
    if denom == 0:
        return "undefined"
    g = gcd(num, denom)
    num, denom = num // g, denom // g
    if denom == 1:
        return str(num)
    if denom < 0:
        num, denom = -num, -denom
    return f"{num}/{denom}"

def generate_question(op_type, min_val, max_val, tier, tier_range):
    """Generate a question and answer based on operation type."""
    a = generate_number(min_val, max_val, tier, tier_range)
    b = generate_number(min_val, max_val, tier, tier_range)

    def pretty_fraction(frac):
        """Return pretty version of a fraction or mixed number."""
        if frac.denominator == 1:
            return str(frac.numerator)
        mixed = divmod(frac.numerator, frac.denominator)
        frac_str = f"{mixed[1]}/{frac.denominator}"
        pretty = fractions_unicode.get(frac_str, frac_str)
        return f"{mixed[0]}{pretty}" if mixed[0] > 0 else pretty
    
    if op_type == "add":
        question = f"{a} + {b}"
        answer = a + b
        return question, str(answer)
    
    elif op_type == "add3":  # New operation for adding three numbers
        c = generate_number(min_val, max_val, tier, tier_range)
        question = f"{a} + {b} + {c}"
        answer = a + b + c
        return question, str(answer)  
      
    elif op_type == "add_dec":  # New operation for adding decimals
        a = generate_number(min_val*10, max_val*10, tier, tier_range) / 10
        b = generate_number(min_val*10, max_val*10, tier, tier_range) / 10
        question = f"{a} + {b}"
        answer = a + b
        return question, f"{answer:.1f}"
    
    elif op_type == "subtract":
        # Ensure non-negative result by making a >= b
        if a < b:
            a, b = b, a  # Swap if a < b
        question = f"{a} - {b}"
        answer = a - b
        return question, str(answer)

    elif op_type == "multiply":
        question = f"{a} × {b}"
        answer = a * b
        return question, str(answer)

    elif op_type == "divide":
        a = generate_number(min_val, max_val / 2, tier, tier_range)
        if a== 0:
            a = 1
        new_max_val = int(max_val // a)  # work out the max product to still fit between our range
        if new_max_val <3:
            new_max_val = 3
        b = generate_number(2, new_max_val, tier, tier_range)
        if(isprime(b)):
            b += 1  # Avoid prime numbers to ensure integer division
        # Ensure integer division result
        product = a * b
        question = f"{product} ÷ {b}"
        answer = a
        return question, str(answer)

    elif op_type == "add_fraction_same_denominator":
        same_denom = [2, 3, 4, 5, 6, 8]
        denom = random.choice(same_denom)
        a_num = random.randint(1, denom - 1)
        b_num = random.randint(1, denom - a_num)  # ensure < 1
        a = Fraction(a_num, denom)
        b = Fraction(b_num, denom)

        answer = a + b
        question = f"{pretty_fraction(a)} + {pretty_fraction(b)}"
        return question, pretty_fraction(answer)

    elif op_type == "add_fraction_different_denominator":
        denoms = [2, 3, 4, 5, 6, 7, 8]
        denom_a, denom_b = random.sample(denoms, 2)  # ensure different denominators
        a_num = random.randint(1, denom_a - 1)
        b_num = random.randint(1, denom_b - 1)
        
        a = Fraction(a_num, denom_a)
        b = Fraction(b_num, denom_b)
        answer = a + b

        question = f"{pretty_fraction(a)} + {pretty_fraction(b)}"
        return question, pretty_fraction(answer)

    elif op_type == "place_value":
        place_names = ["ones", "tens", "hundreds", "thousands", "ten-thousands", "hundred-thousands", "millions", "ten-millions", "hundred-millions", "billions"]
        digits = list(str(a))
        output = []
        length = len(digits)

        for i, digit in enumerate(digits):
            if digit != '0':
                place_index = length - i - 1
                place = place_names[place_index] if place_index < len(place_names) else f"10^{place_index}"
                output.append(f"{digit} {place}")

        question = f"<font size=75% color=dark_gray>What is</font>\n{' + '.join(output)}?"
        answer = a 
        return question, str(answer)

    elif op_type == "place_value_reverse":
        place_names = ["ones", "tens", "hundreds", "thousands", "ten-thousands", "hundred-thousands", 
                    "millions", "ten-millions", "hundred-millions", "billions"]

        # Create a number with non-repeating digits
        while True:
            num_digits = len(str(a)) # Use the length of the number generated through the common code as the number of digits
            digits = random.sample("123456789", 1) + random.sample("0123456789", num_digits - 1)
            random.shuffle(digits)
            if len(set(digits)) == len(digits):  # Ensure uniqueness
                break

        number = int("".join(digits))
        digit_list = list(str(number))
        length = len(digit_list)

        # Randomly pick a digit to ask about
        index = random.randint(0, length - 1)
        digit = int(digit_list[index])
        place_index = length - index - 1

        place = place_names[place_index] if place_index < len(place_names) else f"10^{place_index}"
        place_value = digit * (10 ** place_index)

        question = f"<font size=75% color=dark_gray>What value is the </font>\n{digit} in {number}?"
        answer = str(place_value)

        return question, answer
    
    elif op_type == "fraction":
        # Generate fraction addition
        denom1 = generate_number(1, max_val // 2, tier, tier_range) or 1
        denom2 = generate_number(1, max_val // 2, tier, tier_range) or 1
        num1 = generate_number(0, max_val // 2, tier, tier_range)
        num2 = generate_number(0, max_val // 2, tier, tier_range)
        question = f"{num1}/{denom1} + {num2}/{denom2}"
        # Compute answer as simplified fraction
        frac = Fraction(num1, denom1) + Fraction(num2, denom2)
        answer = simplify_fraction(frac.numerator, frac.denominator)
        return question, answer
    return None, None

def generate_question_set(question_type, level):
    """Generate ordered question/answer pairs for a given level."""
    operations = question_type.get(level, {})
    if not operations:
        return []

    # Determine the maximum tier across all operations
    max_tier = max(max(config["tiers"]) for op_configs in operations.values() for config in op_configs)
    questions_by_tier = [[] for _ in range(max_tier + 1)]
    seen_questions = {}  # Dictionary to track questions by op_type, config index, and tier

    for op_type, configs in operations.items():
        # Initialize seen_questions for this op_type if not present
        if op_type not in seen_questions:
            seen_questions[op_type] = {}
        for config_idx, config in enumerate(configs):
            qty = config["qty"]
            min_val = config["min"]
            max_val = config["max"]
            tiers = config["tiers"]
            fontsize = config.get("fontsize", 120)  # Default font size if not specified
            total_tiers = len(tiers)
            if total_tiers == 0:
                continue

            # Initialize seen_questions for this config_idx and tiers
            seen_questions[op_type][config_idx] = {tier: set() for tier in tiers}

            # Distribute questions across tiers
            questions_per_tier = qty // total_tiers
            extra_questions = qty % total_tiers

            for tier in tiers:
                num_questions = questions_per_tier + (1 if tier - tiers[0] < extra_questions else 0)
                for _ in range(num_questions):
                    max_attempts = 100
                    for attempt in range(max_attempts):
                        question, answer = generate_question(op_type, min_val, max_val, tier, tiers)
                        if question and answer:
                            # Check if question is unique for this op_type, config, and tier
                            if question not in seen_questions[op_type][config_idx][tier]:
                                seen_questions[op_type][config_idx][tier].add(question)
                                questions_by_tier[tier].append((question, answer, tier, fontsize))
                                break
                            # If duplicate, try again
                            if attempt == max_attempts - 1:
                                print(f"Warning: Max attempts reached for {op_type} config {config_idx} in tier {tier}; accepting duplicate: {question}")
                                questions_by_tier[tier].append((question, answer, tier, fontsize))
                                break

    # Shuffle questions within each tier and combine
    result = []
    for tier in range(1, max_tier + 1):
        random.shuffle(questions_by_tier[tier])
        result.extend(questions_by_tier[tier])

    return result

def set_slide_background(slide, red, green, blue):
    # Set background color
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(red, green, blue)

def main():
    # Generate questions for each level
    for level in question_type:
        if level != 2:
            continue

        print(f"\nLevel {level} Questions:")
        
        prev_tier = 0
        question_number = 0
        # Create presentation
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)

        # Set margins and textbox dimensions for the Question
        left = Inches(0.2)
        top = Inches(0.5)
        width = prs.slide_width - Inches(0.4)  # Leave small margin on sides
        height = Inches(3.5)  # Adjust height as needed

        # --- Year Level Slide ---
        slide_q = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide
        q_box = add_textbox(slide_q, f"YEAR {level}", left=left, top=Inches(2), width=width, height=Inches(2), default_font_size=120)
        q_box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)  # Black text

        questions = generate_question_set(question_type, level)
        for i, (question, answer, tier, fontsize) in enumerate(questions, 1):
            print(f"T{tier}Q{i}: {question} = {answer}")
            if tier != prev_tier:
                tier_count = sum(1 for _, _, t, _ in questions if t == tier)
                # --- Create a Tier slide ---
                slide_t = prs.slides.add_slide(prs.slide_layouts[6])  
                set_slide_background(slide_t, 250, 210, 180)
                add_textbox(slide_t, f"<font size=40%>YEAR {level}</font>\nRound {tier}\n\n<font size=60%>{tier_count} Questions</font>", left=left, top=Inches(1), width=width, height=Inches(2), default_font_size=100)
                prev_tier = tier
                question_number = 0

            question_number += 1
            # --- Question Slide ---
            slide_q = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide
            # Footer text
            add_textbox(slide_q, f"QUESTION {question_number} of {tier_count}", Inches(0.2), prs.slide_height-Inches(1), Inches(3), Inches(1), default_font_size=18, align=PP_ALIGN.LEFT, default_color=COLOR_MAP['dark_gray'])
            # Question Text ---
            q_box = add_textbox(slide_q, question, left, top, width, height, default_font_size=fontsize)
            # Tier information
            t_box = add_textbox(slide_q, f"Round {tier}", prs.slide_width-Inches(3), prs.slide_height-Inches(1), Inches(3), Inches(1), default_font_size=18, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.BOTTOM)
            t_box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(180, 180, 180)  # Gray text

            # --- Answer Slide ---
            slide_a = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide
            # Set background color
            set_slide_background(slide_a, 155, 255, 200)
            # Show the question in small text at the top
            clean_question = question.replace('\n', ' ')
            add_textbox(slide_a, f"Q. {clean_question}", left=left, top=Inches(0.3), width=width, height=Inches(1), default_font_size=38)
            # Show the answer
            a_box = add_textbox(slide_a, answer, left=left, top=Inches(2.5), width=width, height=Inches(2), default_font_size=120)
            a_box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 176, 80)  # Green text
            # Show the question #
            add_textbox(slide_a, f"Answer {question_number} of {tier_count}", 
                        Inches(0.2), prs.slide_height-Inches(1), Inches(3), Inches(1), 
                        default_font_size=18, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.BOTTOM, default_color=COLOR_MAP['dark_gray'])
            # Show the tier
            t_box = add_textbox(slide_a, f"Round {tier}", 
                                prs.slide_width-Inches(3), prs.slide_height-Inches(1), Inches(3), Inches(1), 
                                default_font_size=18, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.BOTTOM, default_color=COLOR_MAP['dark_gray'])
            # t_box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(180, 180, 180)  # Gray text

        # Save presentation
        pptx_path = f"c:/temp/data/Maths_Competition_Yr{level}.pptx"
        prs.save(pptx_path)


if __name__ == "__main__":
    random.seed()  # Initialize random seed for reproducibility
    main()