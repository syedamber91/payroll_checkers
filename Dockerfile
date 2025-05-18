FROM python:3.9-slim


# Install Java using default-jre
RUN apt-get update && \
    apt-get install -y default-jre && \
    apt-get install -y default-jre mime-support && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy entire project structure as-is
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Set environment variables
ENV JAVA_HOME=/usr/lib/jvm/java-11-openjdk-amd64
ENV PATH="${JAVA_HOME}/bin:${PATH}"

# Verify Java and JAR location
RUN java -version && \
    ls -la xls-xlsx-converter/target/*.jar

EXPOSE 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]