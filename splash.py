import matplotlib.pyplot as plt
import matplotlib.animation as animation

# Create a figure and a set of axes
fig, ax = plt.subplots()

# Create a line object
line, = ax.plot([], [])

# Initialize the data
x = []
y = []

# This function will be called repeatedly to update the plot
def animate(i):
    # Update the data
    x.append(i)
    y.append(i**2)

    # Set the line data
    line.set_data(x, y)
    print(x,y)
    # Return the line object
    return line

# Create the animation
ani = animation.FuncAnimation(fig, animate, interval=100, save_count=1000)

# Show the plot
plt.show()