# IBM AI Engineering Project Work
# creator: Benjamin Pachter
# date: 6/10/2025
#
# NN1HiddenLayer.py
# directive: demonstrate the basic syntax of a PyTorch neural network with one hidden layer
import torch
import torch.nn as nn
import torchvision.transforms as transformes
import torchvision.datasets as dsets
import torch.nn.functional as F
import matplotlib.pylab as plt
import numpy as np

# define a function to plot acuracy and loss:
def plot_accuracy_loss(training_results):
    plt.subplot(2,1,1)
    plt.plot(training_results['training_loss'], 'r')
    plt.ylabel('loss')
    plt.title('training loss iterations')
    plt.subplot(2,1,2)
    plt.plot(training_results['validation_accuracy'])
    plt.ylabel('accuracy')
    plt.xlabel('epochs')
    plt.show()

# define a function to plot the model parameters
def print_model_paramemters(model):
    count = 0
    for ele in model.state_dict():
        count += 1
        if count % 2 != 0:
            print("The following are the parameters for the layer ", count // 2+1)
        if ele.find("bias") != -1:
            print("The size of the bias is: ", model.state_dict()[ele].size())
        else:
            print("This size of the weights is:", model.state_dict()[ele].size())

# function to display data
def show_data(data_sample):
    plt.imshow((data_sample.numpy().reshape(28, 28), cmap='gray'))
    plt.show()        


### model
# define neural network class
class Net(nn.Module):
    # Constructor
    def __init__(self, D_in, H, D_out):
        super(Net, self).__init__()
        self.linear1 = nn.Linear(D_in, H)
        self.linear2 = nn.Linear(H, D_out)

    # Prediction
    def forward(self, x):
        x = torch.sigmoid(self.linear1(x))
        x = self.linear2(x)