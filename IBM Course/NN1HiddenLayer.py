# IBM AI Engineering Project Work
# creator: Benjamin Pachter
# date: 6/10/2025
#
# NN1HiddenLayer.py
# directive: demonstrate the basic syntax of a PyTorch neural network with one hidden layer
import torch
import torch.nn as nn
import torch.utils.data.dataloader
import torchvision.transforms as transforms
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
def print_model_parameters(model):
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
    plt.imshow(data_sample.numpy().reshape(28, 28), cmap='gray')
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

# Model training function
def train(model, criterion, train_loader, validation_loader, optimizer, epochs=100):
    i = 0
    useful_stuff = {'training_loss': [],'validation_accuracy': []}  
    for epoch in range(epochs):
        for i, (x, y) in enumerate(train_loader): 
            optimizer.zero_grad()
            z = model(x.view(-1, 28 * 28))
            loss = criterion(z, y)
            loss.backward()
            optimizer.step()
            # loss for every iteration
            useful_stuff['training_loss'].append(loss.data.item())
        correct = 0
        for x, y in validation_loader:
            # validation 
            z = model(x.view(-1, 28 * 28))
            _, label = torch.max(z, 1)
            correct += (label == y).sum().item()
        accuracy = 100 * (correct / len(validation_dataset))
        useful_stuff['validation_accuracy'].append(accuracy)
    return useful_stuff

# create training dataset from MNIST
train_dataset = dsets.MNIST(root='./data', train=True, download=True, transform=transforms.ToTensor())

# create validating dataset
validation_dataset = dsets.MNIST(root='./data', downnload=True, transform=transforms.ToTensor())

# create criterion function
criterion = nn.CrossEntropyLoss()

# create Data Loader for both train dataset and validate dataset
train_loader = torch.utils.data.dataloader(dataset=train_dataset, batch_size = 2000, shuffle = True)
validation_loader = torch.utils.data.dataloader(dataset=validation_dataset, batch_size = 5000, shuffle = False)


# create the model with 100 neurons
input_dim = 28 * 28
hidden_dim = 100
output_dim = 10

model = Net(input_dim, hidden_dim, output_dim)

# print model parameters
print_model_parameters(model)


# set the learning rate annd optimizer
learning_rate = 0.01
optimizer = torch.optim.SGD(model.parameters(), lr=learning_rate)


# EPOCHS
# train the model using epochs
training_results = train(model, criterion, train_loader, validation_loader, optimizer, epochs = 30)


# Analysis
# plot accuracy and loss
plot_accuracy_loss(training_results) 

# plot first five misclassified samples
count = 0
for x, y in validation_dataset:
    z = model(x.reshape(-1, 28 * 28))
    _,yhat = torch.max(z, 1)
    if yhat != y:
        show_data(x)
        count += 1
    if count >= 5:
        break

    # Testing another push to Icarus

    
