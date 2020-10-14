var express = require('express')
import { Request, Response, NextFunction } from 'express'
import cors from 'cors'
import multer from 'multer'
var winax = require('winax');
const word = new winax.Object('Word.Application');
//https://docs.microsoft.com/en-us/javascript/api/word/word.application

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, `${__dirname}/../upload`)
    },
    filename: function (req, file, cb) {
        cb(null, file.fieldname + '-' + Date.now() + '.pdf')
    }
})

const upload = multer({
    // dest: `${__dirname}/../upload`,
    storage: storage,
    fileFilter: (req:Request, file, cb) => {
        // if (file.mimetype == "image/png" || file.mimetype == "image/jpg" || file.mimetype == "image/jpeg") {
        if (file.mimetype == "application/pdf") {
            cb(null, true);
        } else {
            cb(null, false);
            return cb(new Error('Only .pdf format allowed!'));
        }
    }
})

var app = express()
app.use(cors())
app.use('/download', express.static('download'))
var port = 9090

app.get('/', async function (req: Request, res: Response) {
    res.send('hello')
})

app.post('/pdf2doc', upload.single('file'), async (req: Request, res: Response) => {
    try {
        const { originalname, mimetype, path } = req.file
        if (mimetype == 'application/pdf') {
            //do
            console.log(path)
            const doc = word.documents.open(path)
            const name = originalname.replace('.pdf', '.docx')
            doc.SaveAs2(`${__dirname}/../download/${name}`, 16)
            doc.close()
            res.json(`http://localhost:9090/download/${name}`);
        } else {
            res.json({
                err: true,
                msg: 'file type is not pdf'
            });
        }
    } catch (err) {
        console.log(err)
        res.sendStatus(400);
    }
})

app.listen(port, '0.0.0.0', () => console.log(`app listening on port ${port}!`))
