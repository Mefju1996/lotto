import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from keras.models import Sequential
from keras.layers import Dense

# Wczytanie danych
data = pd.read_excel('C:/Users/Klaudia/Desktop/wyniki_lotto.ods', engine='odf', sheet_name='Arkusz1')

# Przygotowanie danych
X = data.iloc[:, :6].values  # Zakładam, że kolumny A-F to liczby losowań
y = data.iloc[:, :6].values  # Zakładam, że przewidujemy te same liczby

# Dodanie cech: suma liczb i różnice
sums = X.sum(axis=1).reshape(-1, 1)
diffs = np.diff(X, axis=1).max(axis=1).reshape(-1, 1)
X = np.hstack((X, sums, diffs))

# Normalizacja danych
scaler = StandardScaler()
X = scaler.fit_transform(X)

# Podział danych
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Budowa modelu
model = Sequential()
model.add(Dense(128, input_dim=X_train.shape[1], activation='relu'))
model.add(Dense(64, activation='relu'))
model.add(Dense(6, activation='linear'))  # Wyjście dla 6 liczb

# Kompilacja modelu
model.compile(loss='mean_squared_error', optimizer='adam', metrics=['mae'])

# Trenowanie modelu
model.fit(X_train, y_train, epochs=100, batch_size=32, verbose=1)

# Ocena modelu
loss, mae = model.evaluate(X_test, y_test)
print(f'Loss: {loss}, MAE: {mae}')

# Generowanie prognoz
predictions = model.predict(X_test)

# Konwersja wyników na liczby całkowite
predictions = np.rint(predictions).astype(int)

# Zapis wyników do pliku tekstowego
output_path = 'C:/Users/Klaudia/Desktop/predictions.txt'
np.savetxt(output_path, predictions, fmt='%d', header='Prognozowane wyniki', comments='')

# Wyświetlenie ścieżki zapisu
print(f'Wyniki zapisano w pliku: {output_path}')
