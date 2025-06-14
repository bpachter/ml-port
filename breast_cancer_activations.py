# Directive: Binary classification on Breast Cancer dataset using a 1-hidden-layer neural network with varying activation functions.
# Author: Benjamin Pachter
# Framework: PyTorch + Scikit-Learn
# Dataset: sklearn.datasets.load_breast_cancer (binary classification)
# Initialized: 6/13/2025
#
#   Parameters:
#       Hidden layer: 30 neurons
#       Epochs: 900
#       Optimizer: Adam (lr = 0.01)
#       Loss: Binary Cross Entropy
#
#   Hypothesis:
#      Using the ReLU activation function in the hidden layer will result in faster convergence and lower training loss compared to Tanh and Sigmoid, due to its non-saturating gradient properties and computational efficiency.
#   Conclusion:
#      See breast_cancer_activations_output.png for results. ReLU achieved the fastest and most stable convergence, minimizing training loss earlier than both tanh and sigmoid.
#      While final accuracy was similar across all functions as expected (~0.97–1.00), sigmoid showed slower and more unstable learning, affirming its known limitations (vanishing gradients).
#      ReLU remains the preferred choice for hidden layers in feedforward networks on real data.

import torch
import torch.nn as nn
import torch.optim as optim
from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score
import matplotlib.pyplot as plt

# load the breast cancer dataset
data = load_breast_cancer()
X, y = data.data, data.target

# standardize features to have mean 0 and std 1 (critical for ANN performance)
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# split data into train and test sets
X_train, X_test, y_train, y_test = train_test_split(
    X_scaled, y, test_size=0.20, random_state=3
)

# convert to torch tensors
X_train = torch.tensor(X_train, dtype=torch.float32)
X_test = torch.tensor(X_test, dtype=torch.float32)
y_train = torch.tensor(y_train, dtype=torch.float32).unsqueeze(1)
y_test = torch.tensor(y_test, dtype=torch.float32).unsqueeze(1)

# define a simple feedforward neural network class
class SimpleNet(nn.Module):
    def __init__(self, input_dim, hidden_dim, activation_fn):
        super(SimpleNet, self).__init__()
        self.linear1 = nn.Linear(input_dim, hidden_dim)
        self.activation = activation_fn
        self.linear2 = nn.Linear(hidden_dim, 1)
        self.output_activation = nn.Sigmoid()  # binary classification

    def forward(self, x):
        x = self.linear1(x)
        x = self.activation(x)
        x = self.linear2(x)
        x = self.output_activation(x)
        return x

# define a training loop: 300 epochs as default
def train_model(model, optimizer, criterion, X_train, y_train, X_test, y_test, epochs=300):
    train_loss_list = []
    test_acc_list = []

    for epoch in range(epochs):
        # forward pass
        y_pred = model(X_train)
        loss = criterion(y_pred, y_train)

        # backward pass
        optimizer.zero_grad()
        loss.backward()
        optimizer.step()

        # record metrics
        with torch.no_grad():
            y_test_pred = model(X_test)
            test_preds = (y_test_pred > 0.5).float()
            acc = accuracy_score(y_test.numpy(), test_preds.numpy())

        train_loss_list.append(loss.item())
        test_acc_list.append(acc)

        if (epoch+1) % 20 == 0:
            print(f"Epoch {epoch+1}: Train Loss={loss.item():.4f}, Test Accuracy={acc:.4f}")

    return train_loss_list, test_acc_list

# training settings
input_dim = X_train.shape[1]
hidden_dim = 30
epochs = 900
learning_rate = 0.01
criterion = nn.BCELoss()

# compare different activation functions
activations = {
    "Sigmoid": torch.sigmoid,
    "Tanh": torch.tanh,
    "ReLU": torch.relu
}

results = {}

for name, act_fn in activations.items():
    print(f"\nTraining with {name} activation")
    model = SimpleNet(input_dim, hidden_dim, act_fn)
    optimizer = optim.Adam(model.parameters(), lr=learning_rate)

    loss_list, acc_list = train_model(
        model, optimizer, criterion, X_train, y_train, X_test, y_test, epochs
    )

    results[name] = {"loss": loss_list, "acc": acc_list}

# plot results
plt.figure(figsize=(14, 6))

# plot loss
plt.subplot(1, 2, 1)
for name in results:
    plt.plot(results[name]["loss"], label=name)
plt.title("Training Loss")
plt.xlabel("Epochs")
plt.ylabel("Loss")
plt.legend()

# plot accuracy
plt.subplot(1, 2, 2)
for name in results:
    plt.plot(results[name]["acc"], label=name)
plt.title("Validation Accuracy")
plt.xlabel("Epochs")
plt.ylabel("Accuracy")
plt.legend()

plt.tight_layout()
plt.show()
