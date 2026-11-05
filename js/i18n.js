// Menus and buttons in Hindi (Settings › Language). Lessons stay in English for now.
import { store } from './store.js';
const HI = {
  'Learn': 'सीखें', 'Practise': 'अभ्यास', 'Track': 'प्रगति', 'My stuff': 'मेरी चीज़ें',
  'Home': 'होम', 'My courses': 'मेरे कोर्स', 'Daily review': 'रोज़ का रिवीज़न', 'Tests': 'टेस्ट', 'Playground': 'प्लेग्राउंड', 'Weekly challenge': 'साप्ताहिक चुनौती', 'Interview prep': 'इंटरव्यू तैयारी',
  'Function library': 'फ़ंक्शन लाइब्रेरी', 'Shortcuts': 'शॉर्टकट', 'Topic index': 'विषय सूची', 'Study plan': 'स्टडी प्लान', 'Journal': 'डायरी', 'Mistakes': 'गलतियाँ', 'Insights': 'आँकड़े', 'Certificates': 'सर्टिफ़िकेट',
  'My notes': 'मेरे नोट्स', 'Saved': 'सेव किए गए', 'Downloads': 'डाउनलोड', 'Help': 'मदद', 'Settings': 'सेटिंग्स', 'Admin': 'एडमिन',
  'Search': 'खोजें', 'Search lessons, topics, functions…': 'पाठ, विषय, फ़ंक्शन खोजें…', 'Student sign-in': 'स्टूडेंट साइन-इन', 'Sign out': 'साइन आउट', 'Map': 'मैप', 'Review': 'रिवीज़न', 'More': 'और',
  'day streak': 'दिन लगातार', 'min today': 'मिनट आज', 'to review': 'रिवीज़न बाकी', 'Fold the side panel': 'साइड पैनल छोटा करें', 'Open the side panel': 'साइड पैनल खोलें',
  'Next': 'आगे', 'Back': 'पीछे', 'Check': 'जाँचें', 'Hint': 'संकेत', 'Continue': 'जारी रखें', 'Start': 'शुरू करें', 'Skip': 'छोड़ें', 'Save': 'सेव करें', 'Close': 'बंद करें', 'Done': 'हो गया',
  'Lecture': 'लेक्चर', 'Demo': 'डेमो', 'Lab': 'लैब', 'Quiz': 'क्विज़', 'Recap': 'सारांश', 'Outline': 'रूपरेखा', 'Full screen': 'फ़ुल स्क्रीन', 'Exit focus': 'फ़ोकस से बाहर', 'Notes': 'नोट्स',
  'Downloads & screenshots': 'डाउनलोड और स्क्रीनशॉट', 'Ask your teacher': 'टीचर से पूछें', 'Show me the answer': 'जवाब दिखाएँ', 'Show solution': 'हल दिखाएँ', 'Up next': 'अगला',
  'Free access': 'मुफ़्त एक्सेस', 'Signed in': 'साइन इन', 'Take the tour': 'टूर देखें', 'Page not found': 'पेज नहीं मिला', 'Go home': 'होम पर जाएँ',
};
export const t = s => (store.d && store.d.settings && store.d.settings.lang === 'hi' && HI[s]) || s;
