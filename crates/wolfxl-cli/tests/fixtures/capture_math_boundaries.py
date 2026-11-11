"""Generate the math/information oracle including floating-point boundaries.

Run with the same arguments as capture_math_information.py.
"""

import capture_math_information as capture


if __name__ == "__main__":
    capture.CASES["ATAN2"].extend(["ATAN2(-1,-0)", "ATAN2(-1,-1E-300)"])
    for function in ["CEILING", "FLOOR"]:
        capture.CASES[function].extend(
            f"{function}({number},{significance})"
            for number, significance in [
                ("1E300", "1E-100"),
                ("-1E300", "1E-100"),
                ("-1E300", "-1E-100"),
                ("1E-300", "1E100"),
                ("-1E-300", "1E100"),
                ("-1E-300", "-1E100"),
            ]
        )
    capture.main()
