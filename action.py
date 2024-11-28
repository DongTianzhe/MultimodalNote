import pyperclip
class Action:
    def __init__(self):
        self.scene = '1'
        self.action = '1'

        self.picturePath = ''

        self.speechText = ''
        
        self.musicScorePath = ''
        self.scoreText = ''
        self.key = 'c'
        self.timeNumeratorText = '4'
        self.timeDenominator = '4'
        self.tempo = ''
        self.musicScoreInstrumentsCount = {'Vocal': 0, 'Guitar': 0, 'Bass': 0, 'Drum': 0, 'Treble clef': 0, 'Alto clef': 0, 'Bass clef': 0}
        
        self.gestureText = ''
        self.facialExpressionText = ''
        self.movementText = ''

        self.shotText = ''
        self.focalLensText = ''
        self.cameraMovementText = ''

        self.transitionText = ''
        self.specialEffectText = ''

    def getLilyPondscript(self):
        contentList = self.scoreText.split('\n')
        currentRow = 0
        head = 'global = {\n\\time' + f'{self.timeNumeratorText}/{self.timeDenominator}\n\key {self.key} \major\n'
        if self.tempo != '':
            head += f'\\tempo {self.timeDenominator} = {self.tempo}\n'
        head += '}\n'
        script = head + '<<\n'
        for inst in self.musicScoreInstrumentsCount:
            if inst == 'Vocal':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new Staff{\n\clef treble\n\global\n' + contentList[currentRow] + '\n}\n'
                    currentRow += 1
            elif inst == 'Guitar':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new StaffGroup <<\n\\new Staff{\n\clef treble\n\global\n\\relative{' + contentList[currentRow] + '}}\n' + '\\new TabStaff{\n\global\\relative{' + contentList[currentRow] + '}\n}\n>>\n'
                    currentRow += 1
            elif inst == 'Bass':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new StaffGroup <<\n\\new Staff{\n\clef bass\n\global\n\\relative{' + contentList[currentRow] + '}}\n' + '\\new TabStaff \with {\nstringTunings = #bass-tuning\n}\global\n\\relative{\n' + contentList[currentRow] + '\n}\n>>\n'
                    currentRow += 1
            elif inst == 'Drum':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new DrumStaff <<\n\global\n\\new DrumVoice {\\voiceOne \drummode{' + contentList[currentRow] + '}\n' + '}}\n\\new DrumVoice {\\voiceTwo \drummode{' + contentList[currentRow + 1] + '}}\n>>\n'
                    currentRow += 2
            elif inst == 'Treble clef':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new Staff{\n\clef treble\n\global\n' + contentList[currentRow] + '\n}\n'
                    currentRow += 1
            elif inst == 'Alto clef':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new Staff{\n\clef alto\n\global\n' + contentList[currentRow] + '\n}\n'
                    currentRow += 1
            elif inst == 'Bass clef':
                for i in range(self.musicScoreInstrumentsCount[inst]):
                    script += '\\new Staff{\n\clef bass\n\global\n' + contentList[currentRow] + '\n}\n'
                    currentRow += 1
        script += '>>'
        pyperclip.copy(script)
    
    def getDramaticActionText(self):
        combinedText = ''
        if self.gestureText != '':
            combinedText += f'Gesture: {self.gestureText}\n'
        if self.facialExpressionText != '':
            combinedText += f'Facial Expression: {self.facialExpressionText}\n'
        if self.movementText != '':
            combinedText += f'Movement: {self.movementText}\n'
        return combinedText
    
    def getFilmingText(self):
        combinedText = ''
        if self.shotText != '':
            combinedText += f'Shot: {self.shotText}\n'
        if self.focalLensText != '':
            combinedText += f'Focal Lens: {self.focalLensText}\n'
        if self.cameraMovementText != '':
            combinedText += f'Camera Movement: {self.cameraMovementText}\n'
        return combinedText
    
    def getEditingText(self):
        combinedText = ''
        if self.transitionText != '':
            combinedText += f'Transition: {self.transitionText}\n'
        if self.specialEffectText != '':
            combinedText += f'Special effect {self.specialEffectText}\n'
        return combinedText

if __name__ == '__main__':
    alpha = Action()
    alpha.musicScoreInstrumentsCount['Guitar'] = 1
    alpha.scoreText = "c' d e f g"
    alpha.getLilyPondscript()