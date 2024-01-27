from collections import Counter

def Average(lst):
    if not lst:
        return 0  # or handle it in a way that makes sense for your application
    return sum(lst) / len(lst)

def calculate_return_rate(sales_cycle):
    actual_lst = []
    ideal_lst = []

    for i in sales_cycle:
        if i > 0:
            actual_lst.append(i)
            ideal_lst.append(i)
        elif i < 0 and ideal_lst:
            ideal_lst[-1] += i

    # Calculate the average of actual quantities using your function
    actual_average = Average(actual_lst)

    # Calculate the average of ideal quantities using your function
    ideal_average = Average(ideal_lst)

    # Single recommendation value (average of corrected values)
    overall_recommendation = actual_average + ideal_average

    return overall_recommendation

# Example usage
print(calculate_return_rate([5.0, 5.0, 5.0, 5.0, 5.0, 4.0, 5.0, 10.0, 7.0, 5.0, 5.0, -5.0]))
