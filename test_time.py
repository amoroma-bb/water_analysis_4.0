#%%
import datetime
today = datetime.date.today()
print(today)
print(str(today))
# %%
arr = [0,1,2,3,4,5,6,7,8]
d = [list(bin(x)[2:]).count('1') for x in arr]
d
# %%
arr = [0,1,2,3,4,5,6,7,8]
sorted(arr,key=lambda x:bin(x).count('1'))

# %%
import tensorflow as tf

# %%
import math
import numpy as np
import matplotlib.pyplot as ply
import tensorflow_datasets as tfds
tfds.disable_progress_bar()
# %%
import logging
logger = tf.get_logger()
logger.setLevel(logging.ERROR)
# %%
dataset, metadata = tfds.load('fashion_mnist', as_supervised=True,with_info=True)
train_dataset, test_dataset = dataset['train'], dataset['test']
# %%
dataset
# %%
train_dataset
# %%
class_names = metadata.features['label'].names
print('Class name: {}'.format(class_names))
# %%
num_train_examples = metadata.splits['train'].num_examples
num_test_examples = metadata.splits['test'].num_examples
print('Number of training exmaples:{}'.format(num_train_examples))
# %%
def normalize(images,labels):
    images = tf.cast(images, tf.float32)
    images /= 255
    return images, labels
train_dataset = train_dataset.map(normalize)
test_dataset = test_dataset.map(normalize)

# %%
train_dataset = train_dataset.cache()
test_dataset = test_dataset.cache()

# %%
for image, label in test_dataset.take(1):
    break
image = image.numpy().reshape((28,28))
ply.figure()
ply.imshow(image, cmap=ply.cm.binary)
ply.colorbar()
ply.grid(False)
ply.show()
# %%
ply.figure(figsize=(10,10))
for i, (image, label) in enumerate(test_dataset.take(25)):
    image = image.numpy().reshape((28,28))
    ply.subplot(5,5,i+1)
    ply.xticks([])
    ply.yticks([])
    ply.grid(False)
    ply.imshow(image, cmap=ply.cm.binary)
    ply.xlabel(class_names[label])
ply.show()
# %%
typed = "aaleex"
ans = ''
for x in typed:
    ans.join(x)
    print(ans)
# %%
mat = [[1,1,0,0,0],[1,1,1,1,0],[1,0,0,0,0],[1,1,0,0,0],[1,1,1,1,1]]
ans =[]
for i in range(len(mat)):
    for m in range(len(mat[i])):
        if mat[i][m] == 0:
            ans.append((i,m))
            break
        elif mat[i][-1] == 1:
            ans.append((i,len(mat[i])))
            break

k = sorted(ans, key=lambda x: x[1])
t = []
for i in range(3):
    t.append(
k
# %%
t = [(2,1),(0,2),(3,2)]
t[0][0]
# %%
