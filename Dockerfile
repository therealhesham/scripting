FROM node:18-bullseye

# تحديث النظام وتثبيت LibreOffice والخطوط العربية الهامة
# لضمان خروج ملفات الـ PDF باللغة العربية بشكل منسق ومقروء
RUN apt-get update && \
    apt-get install -y \
    libreoffice \
    fonts-hosny-amiri \
    fonts-kacst \
    fonts-sil-scheherazade \
    fonts-arabeyes \
    && rm -rf /var/lib/apt/lists/*

# إعداد مسار العمل داخل الحاوية
WORKDIR /usr/src/app

# نسخ ملفات تعريف المشروع وتثبيت الحزم
COPY package*.json ./
RUN npm install

# نسخ باقي ملفات المشروع
COPY . .

# فتح بورت 3000
EXPOSE 3172

# أمر التشغيل الأساسي
CMD [ "node", "server.js" ]
