pipeline {
    agent any
    tools{
        maven 'maven-3.3'
        jdk 'java-1.8'
    }
    stages {
        stage ('Initialize') {  
            steps{
                bat '''
                echo "PATH = %PATH%"
                echo "M2_HOME= %M2_HOME%"
                '''
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
