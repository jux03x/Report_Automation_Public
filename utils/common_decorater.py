# Here you can find a number of common decorater, which can be used to make your code look clean.
# Also found in libaries like logging, functools.wraps, Loguru, structlog, etc.
# For Basic scripts, use these labaries. They are far more advanced.
# For me this script helps understand, how to write and use decoraters.
# It also can be used for project specific decoraters.
# It also gives me the freedom to collect usefull decoraters which aren't found in any libaries.
# It also gives me the freedom to choose selfe written code and local scripts over installt libaries.
import time


# Mark your function end, feel free to select length and symbol which could mark your ending
len_of_beg = 100
sym_of_beg = "."

def begin_of_func(func):
    global len_of_beg
    global sym_of_beg
    def beginning(*args, **kwargs):
        print("\n #Start", sym_of_beg * len_of_beg, "\n")
        func(*args, **kwargs)
    return(beginning)


# Mark your function end, feel free to select length and symbol which could mark your ending
len_of_end = 80
sym_of_end = "#"

def end_of_func(func):
    global len_of_end
    global sym_of_end
    def ending(*args, **kwargs):
        func(*args, **kwargs)
        print("\n", sym_of_end * len_of_end, "\n")
    return(ending)

# Wrap your function, feel free to select length and symbol which could mark your ending
len_of_wrapper = 100
sym_of_wrapper = "<>"

def wrapper_of_func(func):
    global len_of_wrapper
    global sym_of_wrapper
    def wrapper(*args, **kwargs):
        print("\n", sym_of_wrapper * len_of_wrapper, "\n")
        func(*args, **kwargs)
        print("\n", sym_of_wrapper * len_of_wrapper, "\n")
    return(wrapper)

# Logger, which logs the fuction input and its output
def logger(func):
    def get_input_output(*args, **kwargs):
        # input
        i1 = f"positional arguments: {args}"
        i2 = f"Key-Words: {kwargs}"

        
        input = str(i1 + " /// " + i2)
        log_out = ("Logged: input: ", input, "// output: ", func(*args, **kwargs), "// Function Name: ", func.__name__)
        print(log_out)
        return log_out
    
    return get_input_output


# Logger, which logs the fuction input and its output in a additional file
def logger_advc(func):
    def get_input_output(*args, **kwargs):
        # input
        l1 = f"positional arguments: {args}"
        l2 = f"Key-Words: {kwargs}"
        input_str = str(l1 + " /// " + l2)

        # Get output
        output = func(*args, **kwargs)

        # Log-String created
        log_entry = (
            f"Logged: input: {input_str} // "
            f"output: {output} // "
            f"Function Name: {func.__name__}\n"
        )

        # Print log in terminal
        print(log_entry)

        # write log in text-file
        with open("function_log.txt", "a", encoding="utf-8") as f:
            f.write(log_entry)

        return output

    return get_input_output  

# Messure Runtime/ Performance
def measure_time(func):
    @wrapper_of_func
    def time_est(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()
        print(f"{func.__name__} took {end - start:.4f} seconds")
        return result
    return time_est

import time

# Maximum of retrys
def retry(times):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for i in range(times):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    print(f"Retry {i+1}/{times}: {e}")
                    time.sleep(1)
            raise Exception("All retries failed")
        return wrapper
    return decorator
