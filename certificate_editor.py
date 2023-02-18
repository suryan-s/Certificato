import cv2
 
# Load the image
img = cv2.imread('page0.png')
 
# Specify the coordinates for the redaction
top_left_x = 531
top_left_y = 297
bottom_right_x = 972
bottom_right_y = 325
x, y, width, height = 531, 297, (bottom_right_x - top_left_x), (bottom_right_y - top_left_y)
 
# Create a red rectangle to cover the desired portion of the image
red = (0, 0, 255)
img[y:y + height, x:x + width] = red
 
# Write text on the red rectangle using a white color
font = cv2.FONT_HERSHEY_SIMPLEX
org = (x + int(width / 4), y + int(height / 2))
fontScale = 1
color = (255, 255, 255)
thickness = 2
text = "xxx@gmail.com"
img = cv2.putText(img, text, org, font, fontScale, color, thickness, cv2.LINE_AA)
 
# Save the resulting image
cv2.imwrite('redacted_image_with_text.png', img)