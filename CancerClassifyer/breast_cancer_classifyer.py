# import libraries
import torch
import pandas as pd
import numpy as np
import torch.nn as nn
import torch.nn.functional as F
import torch.optim as optim
import matplotlib.pyplot as plt
from sklearn.metrics import classification_report, accuracy_score, confusion_matrix


from sklearn.datasets import fetch_openml
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from torch.utils.data import TensorDataset, DataLoader

# load the dataset from OpenML
raw = fetch_openml(name="BreastCancer", version=1, as_frame=True)
df = raw.frame

# map labels to binary: malignant = 1, benign = 0
df['Class'] = df['Class'].map({'malignant': 1, 'benign': 0})

# drop any non-numeric columns (e.g., ID or missing entries)
df = df.select_dtypes(include=np.number).dropna()

# separate input features and target
X = df.drop(columns=['Class'])
y = df['Class']

# split into training and testing (80% train, 20% test)
X_train, X_test, y_train, y_test = train_test_split(
X, y, test_size=0.2, stratify=y, random_state=42
)

# standardize the feature values to zero mean and unit variance
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# convert to PyTorch tensors
X_train_tensor = torch.tensor(X_train_scaled, dtype=torch.float32)
X_test_tensor = torch.tensor(X_test_scaled, dtype=torch.float32)
y_train_tensor = torch.tensor(y_train.values, dtype=torch.long)
y_test_tensor = torch.tensor(y_test.values, dtype=torch.long)

# wrap data in TensorDatasets and DataLoaders
train_dataset = TensorDataset(X_train_tensor, y_train_tensor)
test_dataset = TensorDataset(X_test_tensor, y_test_tensor)

train_loader = DataLoader(train_dataset, batch_size=16, shuffle=True)
test_loader = DataLoader(test_dataset, batch_size=16, shuffle=False)
# define a neural network for binary classification
class ClassificationNet(nn.Module):
    def __init__(self, input_units=30, hidden_units=64, output_units=2):
        super(ClassificationNet, self).__init__()
        
        # first layer: input -> hidden
        self.fc1 = nn.Linear(input_units, hidden_units)
        
        # second layer: hidden -> output
        self.fc2 = nn.Linear(hidden_units, output_units)
    
    def forward(self, x):
        # apply ReLU activation to the hidden layer
        x = F.relu(self.fc1(x))
        
        # raw output logits for CrossEntropyLoss (no softmax here)
        x = self.fc2(x)
        return x

# instantiate the model
model = ClassificationNet(input_units=X_train.shape[1], hidden_units=64, output_units=2)

# print architecture summary
print(model)

# use cross-entropy loss for classification (expects raw logits)
criterion = nn.CrossEntropyLoss()

# use Adam optimizer (you can later try SGD, RMSprop, etc.)
optimizer = optim.Adam(model.parameters(), lr=0.001)

# set training parameters
epochs = 10
train_losses = []
test_losses = []

# training loop
for epoch in range(epochs):
    model.train()  # set model to training mode
    running_loss = 0.0
    
    for X_batch, y_batch in train_loader:
        optimizer.zero_grad()              # clear old gradients
        outputs = model(X_batch)           # forward pass
        loss = criterion(outputs, y_batch) # compute loss
        loss.backward()                    # backpropagation
        optimizer.step()                   # update model parameters
        
        running_loss += loss.item()        # accumulate batch loss
    
    train_loss = running_loss / len(train_loader)
    train_losses.append(train_loss)
    
    # evaluation loop (test set)
    model.eval()  # set model to evaluation mode
    test_loss = 0.0
    with torch.no_grad():
        for X_batch, y_batch in test_loader:
            outputs = model(X_batch)
            loss = criterion(outputs, y_batch)
            test_loss += loss.item()
    
    test_loss /= len(test_loader)
    test_losses.append(test_loss)

    print(f"Epoch [{epoch+1}/{epochs}], Train Loss: {train_loss:.4f}, Test Loss: {test_loss:.4f}")



plt.figure(figsize=(10, 6))
plt.plot(range(1, epochs + 1), train_losses, label='Training Loss')
plt.plot(range(1, epochs + 1), test_losses, label='Test Loss', linestyle='--')
plt.xlabel('Epoch')
plt.ylabel('Loss')
plt.title('Training and Test Loss Curve')
plt.legend()
plt.grid(True)
plt.show()

