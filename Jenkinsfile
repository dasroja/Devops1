pipeline {
    agent any
    tools{
        maven 'maven-3.3'
        jdk 'java-1.8'
    }
    stages {
        stage ('Initialize') {  
            steps{
                bat 'set M2_HOME=C:\devops\softwares\maven\apache-maven-3.5.2'
                bat "set PATH =C:\devops\softwares\maven\apache-maven-3.3.3:%PATH%"
                bat 'mvn -version'
                bat 'mvn clean compile'
        }
        }

        stage ('build') {
            steps{
                bat 'mvn install'
        }
        post{
            success{
                junit"Devops1/target/surefire-reports/*.xml"
                
            }
        }
      }
    }      
}
