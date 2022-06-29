{
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# Cherry Blossoms!\n",
    "\n",
    "If we travel back in time, [cherry blossoms](https://en.wikipedia.org/wiki/Cherry_blossom) were once in full bloom! We don't live in Japan or DC, but in non-COVID times we also have the [Brooklyn Botanic Garden's annual festival](https://www.bbg.org/visit/event/sakura_matsuri_2020).\n",
    "\n",
    "We'll have to make up for it with data-driven cherry blossoms instead. Once upon a time [Data is Plural](https://tinyletter.com/data-is-plural) linked to [a dataset](http://atmenv.envi.osakafu-u.ac.jp/aono/kyophenotemp4/) about when the cherry trees blossom each year. It's completely out of date, but it's quirky in a real nice way so we're sticking with it.\n",
    "\n",
    "## 0. Do all of your importing/setup stuff"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 1,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Requirement already satisfied: xlrd in /Users/amandachen/.pyenv/versions/3.10.3/lib/python3.10/site-packages (2.0.1)\n",
      "\u001b[33mWARNING: You are using pip version 22.0.4; however, version 22.1.2 is available.\n",
      "You should consider upgrading via the '/Users/amandachen/.pyenv/versions/3.10.3/bin/python3.10 -m pip install --upgrade pip' command.\u001b[0m\u001b[33m\n",
      "\u001b[0m"
     ]
    }
   ],
   "source": [
    "!pip install xlrd\n",
    "import pandas as pd\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 1. Read in the file using pandas, and look at the first five rows\n",
    "\n",
    "* *Tip: You will probably need to pip install something to make this Excel file work!*"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 2,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>AD</th>\n",
       "      <th>Full-flowering date (DOY)</th>\n",
       "      <th>Full-flowering date</th>\n",
       "      <th>Source code</th>\n",
       "      <th>Data type code</th>\n",
       "      <th>Reference Name</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>0</th>\n",
       "      <td>801</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1</th>\n",
       "      <td>802</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>2</th>\n",
       "      <td>803</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>3</th>\n",
       "      <td>804</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>4</th>\n",
       "      <td>805</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>...</th>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1210</th>\n",
       "      <td>2011</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1211</th>\n",
       "      <td>2012</td>\n",
       "      <td>101.0</td>\n",
       "      <td>410</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1212</th>\n",
       "      <td>2013</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1213</th>\n",
       "      <td>2014</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1214</th>\n",
       "      <td>2015</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "<p>1215 rows × 6 columns</p>\n",
       "</div>"
      ],
      "text/plain": [
       "        AD  Full-flowering date (DOY) Full-flowering date  Source code  \\\n",
       "0      801                        NaN                 NaN          NaN   \n",
       "1      802                        NaN                 NaN          NaN   \n",
       "2      803                        NaN                 NaN          NaN   \n",
       "3      804                        NaN                 NaN          NaN   \n",
       "4      805                        NaN                 NaN          NaN   \n",
       "...    ...                        ...                 ...          ...   \n",
       "1210  2011                       99.0                 409          8.0   \n",
       "1211  2012                      101.0                 410          8.0   \n",
       "1212  2013                       93.0                 403          8.0   \n",
       "1213  2014                       94.0                 404          8.0   \n",
       "1214  2015                       93.0                 403          8.0   \n",
       "\n",
       "      Data type code          Reference Name  \n",
       "0                NaN                     NaN  \n",
       "1                NaN                     NaN  \n",
       "2                NaN                     NaN  \n",
       "3                NaN                     NaN  \n",
       "4                NaN                     NaN  \n",
       "...              ...                     ...  \n",
       "1210             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1211             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1212             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1213             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1214             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "\n",
       "[1215 rows x 6 columns]"
      ]
     },
     "execution_count": 2,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df = pd.read_excel(\"KyotoFullFlower7.xls\",skiprows=25, na_values=\"-\", dtype={'Full-flowering date':str})\n",
    "df"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 2. Read in the file using pandas CORRECTLY, and look at the first five rows\n",
    "\n",
    "Hrm, how do your column names look? Read the file in again but this time add **a parameter to make sure your columns look right**. How can you tell pandas to skip rows?\n",
    "\n",
    "**TIP: The first year should be 801 AD, and it should not have any dates or anything.**"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>AD</th>\n",
       "      <th>Full-flowering date (DOY)</th>\n",
       "      <th>Full-flowering date</th>\n",
       "      <th>Source code</th>\n",
       "      <th>Data type code</th>\n",
       "      <th>Reference Name</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>0</th>\n",
       "      <td>801</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1</th>\n",
       "      <td>802</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>2</th>\n",
       "      <td>803</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>3</th>\n",
       "      <td>804</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>4</th>\n",
       "      <td>805</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "</div>"
      ],
      "text/plain": [
       "    AD  Full-flowering date (DOY) Full-flowering date  Source code  \\\n",
       "0  801                        NaN                 NaN          NaN   \n",
       "1  802                        NaN                 NaN          NaN   \n",
       "2  803                        NaN                 NaN          NaN   \n",
       "3  804                        NaN                 NaN          NaN   \n",
       "4  805                        NaN                 NaN          NaN   \n",
       "\n",
       "   Data type code Reference Name  \n",
       "0             NaN            NaN  \n",
       "1             NaN            NaN  \n",
       "2             NaN            NaN  \n",
       "3             NaN            NaN  \n",
       "4             NaN            NaN  "
      ]
     },
     "execution_count": 3,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.head(5)"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 3. Look at the final five rows of the data"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 4,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>AD</th>\n",
       "      <th>Full-flowering date (DOY)</th>\n",
       "      <th>Full-flowering date</th>\n",
       "      <th>Source code</th>\n",
       "      <th>Data type code</th>\n",
       "      <th>Reference Name</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>1210</th>\n",
       "      <td>2011</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1211</th>\n",
       "      <td>2012</td>\n",
       "      <td>101.0</td>\n",
       "      <td>410</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1212</th>\n",
       "      <td>2013</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1213</th>\n",
       "      <td>2014</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1214</th>\n",
       "      <td>2015</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "</div>"
      ],
      "text/plain": [
       "        AD  Full-flowering date (DOY) Full-flowering date  Source code  \\\n",
       "1210  2011                       99.0                 409          8.0   \n",
       "1211  2012                      101.0                 410          8.0   \n",
       "1212  2013                       93.0                 403          8.0   \n",
       "1213  2014                       94.0                 404          8.0   \n",
       "1214  2015                       93.0                 403          8.0   \n",
       "\n",
       "      Data type code          Reference Name  \n",
       "1210             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1211             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1212             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1213             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1214             0.0  NEWS-PAPER(ARASHIYAMA)  "
      ]
     },
     "execution_count": 4,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.tail(5)"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 4. Add some NaN values"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "It looks like you should have NaN/missing values in the beginning of the dataset under \"Reference name.\" Read in the file *one more time*, this time making sure all of those missing reference names actually show up as `NaN` instead of `-`.\n",
    "\n",
    "* *Tip: it's another open with reading in the file!*"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 5,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>AD</th>\n",
       "      <th>Full-flowering date (DOY)</th>\n",
       "      <th>Full-flowering date</th>\n",
       "      <th>Source code</th>\n",
       "      <th>Data type code</th>\n",
       "      <th>Reference Name</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>0</th>\n",
       "      <td>801</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1</th>\n",
       "      <td>802</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>2</th>\n",
       "      <td>803</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>3</th>\n",
       "      <td>804</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>4</th>\n",
       "      <td>805</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "      <td>NaN</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>...</th>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1210</th>\n",
       "      <td>2011</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1211</th>\n",
       "      <td>2012</td>\n",
       "      <td>101.0</td>\n",
       "      <td>410</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1212</th>\n",
       "      <td>2013</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1213</th>\n",
       "      <td>2014</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1214</th>\n",
       "      <td>2015</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "<p>1215 rows × 6 columns</p>\n",
       "</div>"
      ],
      "text/plain": [
       "        AD  Full-flowering date (DOY) Full-flowering date  Source code  \\\n",
       "0      801                        NaN                 NaN          NaN   \n",
       "1      802                        NaN                 NaN          NaN   \n",
       "2      803                        NaN                 NaN          NaN   \n",
       "3      804                        NaN                 NaN          NaN   \n",
       "4      805                        NaN                 NaN          NaN   \n",
       "...    ...                        ...                 ...          ...   \n",
       "1210  2011                       99.0                 409          8.0   \n",
       "1211  2012                      101.0                 410          8.0   \n",
       "1212  2013                       93.0                 403          8.0   \n",
       "1213  2014                       94.0                 404          8.0   \n",
       "1214  2015                       93.0                 403          8.0   \n",
       "\n",
       "      Data type code          Reference Name  \n",
       "0                NaN                     NaN  \n",
       "1                NaN                     NaN  \n",
       "2                NaN                     NaN  \n",
       "3                NaN                     NaN  \n",
       "4                NaN                     NaN  \n",
       "...              ...                     ...  \n",
       "1210             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1211             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1212             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1213             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1214             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "\n",
       "[1215 rows x 6 columns]"
      ]
     },
     "execution_count": 5,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 5. What reference is the most commonly used when figuring out cherry blossom flowering dates?\n",
    "\n",
    "If the first result is `\"-\"`, you need to redo the last question."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 6,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "<class 'pandas.core.frame.DataFrame'>\n",
      "RangeIndex: 1215 entries, 0 to 1214\n",
      "Data columns (total 6 columns):\n",
      " #   Column                     Non-Null Count  Dtype  \n",
      "---  ------                     --------------  -----  \n",
      " 0   AD                         1215 non-null   int64  \n",
      " 1   Full-flowering date (DOY)  827 non-null    float64\n",
      " 2   Full-flowering date        827 non-null    object \n",
      " 3   Source code                827 non-null    float64\n",
      " 4   Data type code             827 non-null    float64\n",
      " 5   Reference Name             825 non-null    object \n",
      "dtypes: float64(3), int64(1), object(2)\n",
      "memory usage: 57.1+ KB\n"
     ]
    }
   ],
   "source": [
    "df.info()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 7,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "NEWS-PAPER(ARASHIYAMA)        94\n",
       "SUGIURAKE-NIKKI               38\n",
       "MYOHOIN-HINAMIKI              32\n",
       "OYUDONONO-UENO-NIKKI          26\n",
       "OYUDONONO-UENO-NIKKI [EDA]    21\n",
       "                              ..\n",
       "GONARAIN-GYOSEI-BASSHO         1\n",
       "MORIMITSU-KOKI                 1\n",
       "SANKAIKI                       1\n",
       "TSUKIMOUDE-WAKASHU(PART-3)     1\n",
       "NEWS-PAPET(DAIGO-JI)           1\n",
       "Name: reference_name, Length: 222, dtype: int64"
      ]
     },
     "execution_count": 7,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.columns = df.columns.str.lower().str.replace(\" \",\"_\")\n",
    "df = df.rename(columns={'full-flowering_date_(doy)': 'full_flowering_date_doy'})\n",
    "df.reference_name.value_counts()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 6. Filter the list to only include columns where the `Full-flowering date (DOY)` is not missing"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 8,
   "metadata": {},
   "outputs": [],
   "source": [
    "#Can you do this with df.query???\n",
    "#df.query(\"full_flowering_date_doy != ''\")\n",
    "#df.query(\"full_flowering_date_doy != isna()\")\n",
    "\n",
    "df = df[df.full_flowering_date_doy.notnull()]\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 6.5 Confirm you now have 827 rows"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 7. Make a histogram of the full-flowering date"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 9,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<AxesSubplot:>"
      ]
     },
     "execution_count": 9,
     "metadata": {},
     "output_type": "execute_result"
    },
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "df.full_flowering_date_doy.hist()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 8. Make another histogram of the full-flowering date, but with 39 bins instead of 10"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 10,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<AxesSubplot:>"
      ]
     },
     "execution_count": 10,
     "metadata": {},
     "output_type": "execute_result"
    },
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "df.full_flowering_date_doy.hist(bins=39)"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 9. What's the average number of days it takes for the flowers to blossom? And how many records do we have?\n",
    "\n",
    "Answer these both with one line of code."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 11,
   "metadata": {},
   "outputs": [],
   "source": [
    "#df.describe()?????"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 10. What's the average days into the year cherry flowers normally blossomed before 1900?\n",
    "\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 12,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "105.20728291316527"
      ]
     },
     "execution_count": 12,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.query(\"ad < 1900\").full_flowering_date_doy.mean()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 11. How about after 1900?"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 13,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "100.3125"
      ]
     },
     "execution_count": 13,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.query(\"ad > 1900\").full_flowering_date_doy.mean()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 12. How many times was our data from a title in Japanese poetry?\n",
    "\n",
    "You'll need to read the documentation inside of the Excel file."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 14,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "3.0    342\n",
       "4.0    250\n",
       "1.0    105\n",
       "5.0     59\n",
       "7.0     40\n",
       "2.0     17\n",
       "6.0      9\n",
       "8.0      5\n",
       "Name: source_code, dtype: int64"
      ]
     },
     "execution_count": 14,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "#when source code = 4 \n",
    "df.source_code.value_counts()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 13. Display the rows where our data was from a title in Japanese poetry"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 15,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>ad</th>\n",
       "      <th>full_flowering_date_doy</th>\n",
       "      <th>full-flowering_date</th>\n",
       "      <th>source_code</th>\n",
       "      <th>data_type_code</th>\n",
       "      <th>reference_name</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>149</th>\n",
       "      <td>950</td>\n",
       "      <td>95.0</td>\n",
       "      <td>405</td>\n",
       "      <td>4.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>DAINIHON-SHIRYO</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>178</th>\n",
       "      <td>979</td>\n",
       "      <td>104.0</td>\n",
       "      <td>414</td>\n",
       "      <td>4.0</td>\n",
       "      <td>3.0</td>\n",
       "      <td>DAINIHON-SHIRYO</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>184</th>\n",
       "      <td>985</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>4.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>SHOYUKI</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>188</th>\n",
       "      <td>989</td>\n",
       "      <td>100.0</td>\n",
       "      <td>410</td>\n",
       "      <td>4.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>SHOYUKI</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>218</th>\n",
       "      <td>1019</td>\n",
       "      <td>98.0</td>\n",
       "      <td>408</td>\n",
       "      <td>4.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>SHOYUKI</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>...</th>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1200</th>\n",
       "      <td>2001</td>\n",
       "      <td>96.0</td>\n",
       "      <td>406</td>\n",
       "      <td>4.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1201</th>\n",
       "      <td>2002</td>\n",
       "      <td>91.0</td>\n",
       "      <td>401</td>\n",
       "      <td>4.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1202</th>\n",
       "      <td>2003</td>\n",
       "      <td>98.0</td>\n",
       "      <td>408</td>\n",
       "      <td>4.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1203</th>\n",
       "      <td>2004</td>\n",
       "      <td>92.0</td>\n",
       "      <td>401</td>\n",
       "      <td>4.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1204</th>\n",
       "      <td>2005</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>4.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "<p>250 rows × 6 columns</p>\n",
       "</div>"
      ],
      "text/plain": [
       "        ad  full_flowering_date_doy full-flowering_date  source_code  \\\n",
       "149    950                     95.0                 405          4.0   \n",
       "178    979                    104.0                 414          4.0   \n",
       "184    985                     94.0                 404          4.0   \n",
       "188    989                    100.0                 410          4.0   \n",
       "218   1019                     98.0                 408          4.0   \n",
       "...    ...                      ...                 ...          ...   \n",
       "1200  2001                     96.0                 406          4.0   \n",
       "1201  2002                     91.0                 401          4.0   \n",
       "1202  2003                     98.0                 408          4.0   \n",
       "1203  2004                     92.0                 401          4.0   \n",
       "1204  2005                     99.0                 409          4.0   \n",
       "\n",
       "      data_type_code          reference_name  \n",
       "149              2.0         DAINIHON-SHIRYO  \n",
       "178              3.0         DAINIHON-SHIRYO  \n",
       "184              2.0                 SHOYUKI  \n",
       "188              2.0                 SHOYUKI  \n",
       "218              2.0                 SHOYUKI  \n",
       "...              ...                     ...  \n",
       "1200             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1201             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1202             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1203             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "1204             0.0  NEWS-PAPER(ARASHIYAMA)  \n",
       "\n",
       "[250 rows x 6 columns]"
      ]
     },
     "execution_count": 15,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.query(\"source_code == 4\")"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 16,
   "metadata": {},
   "outputs": [],
   "source": [
    "df = df.rename(columns={'full-flowering_date': 'full_flowering_date'})\n",
    "\n",
    "\n",
    "#df['full-flowering_date'].astype(int)\n",
    "\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 14. Graph the full-flowering date (DOY) over time"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 17,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<AxesSubplot:>"
      ]
     },
     "execution_count": 17,
     "metadata": {},
     "output_type": "execute_result"
    },
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "df.full_flowering_date_doy.plot()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 15. Smooth out the graph\n",
    "\n",
    "It's so jagged! You can use `df.rolling` to calculate a rolling average.\n",
    "\n",
    "The following code calculates a **10-year mean**, using the `AD` column as the anchor. If there aren't 20 samples to work with in a row, it'll accept down to 5. Neat, right?\n",
    "\n",
    "(We're only looking at the final 5)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 18,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "1210    95.9\n",
       "1211    96.9\n",
       "1212    96.4\n",
       "1213    96.6\n",
       "1214    96.0\n",
       "Name: full_flowering_date_doy, dtype: float64"
      ]
     },
     "execution_count": 18,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.rolling(10, on='ad', min_periods=5)['full_flowering_date_doy'].mean().tail()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 19,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<AxesSubplot:>"
      ]
     },
     "execution_count": 19,
     "metadata": {},
     "output_type": "execute_result"
    },
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "df.rolling(10, on='ad', min_periods=5)['full_flowering_date_doy'].mean().plot()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Use the code above to create a new column called `rolling_date` in our dataset. It should be the 20-year rolling average of the flowering date. Then plot it, with the year on the x axis and the day of the year on the y axis.\n",
    "\n",
    "Try adding `ylim=(80, 120)` to your `.plot` command to make things look a little less dire."
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 16. Add a month column\n",
    "\n",
    "Right now the \"Full-flowering date\" column is pretty rough. It uses numbers like '402' to mean \"April 2nd\" and \"416\" to mean \"April 16th.\" Let's make a column to explain what month it happened in.\n",
    "\n",
    "* Every row that happened in April should have 'April' in the `month` column.\n",
    "* Every row that happened in March should have 'March' as the `month` column.\n",
    "* Every row that happened in May should have 'May' as the `month` column.\n",
    "\n",
    "There are **at least two ways to do this.**\n",
    "\n",
    "#### WAY ONE: The bad-yet-simple way\n",
    "\n",
    "If you don't want to use `pd.to_datetime`, you can use this as an sample for updating March. It finds everything with a date less than 400 and assigns `March` to the `month` column:\n",
    "\n",
    "```python\n",
    "df.loc[df['Full-flowering date'] < 400, 'month'] = 'March'\n",
    "```\n",
    "\n",
    "#### WAY TWO: The good-yet-complicated way\n",
    "\n",
    "* When you use `pd.to_datetime`, if pandas doesn't figure it out automatically you can also pass a `format=` argument that explains what the format is of the datetime. You use [the codes here](https://strftime.org/) to mark out where the days, months, etc are. For example, `2020-04-09` would be converted using `pd.to_datetime(df.colname, \"format='%Y-%m-%d\")`.\n",
    "* `errors='coerce'` will return `NaN` for missing values. By default it just yells \"I don't know what to do!!!\"\n",
    "* And remember how we used `df.date_column.dt.month` to get the number of the month? For the name, you use `dt.strftime` (string-formatted-time), and pass it [the same codes](https://strftime.org/) to tell it what to do. For example, `df.date_column.dt.strftime(\"%Y-%m-%d\")` would give you `\"2020-04-09\"`."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 20,
   "metadata": {},
   "outputs": [],
   "source": [
    "#to convert to date time - make column a str\n",
    "#df.full_flowering_date = df.full_flowering_date.astype(int)\n",
    "#df.dtypes\n",
    "\n",
    "df['date_time'] = pd.to_datetime(df.full_flowering_date, format=\"%m%d\")\n",
    "\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 21,
   "metadata": {},
   "outputs": [],
   "source": [
    "df['month'] = df.date_time.dt.month_name()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 22,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>ad</th>\n",
       "      <th>full_flowering_date_doy</th>\n",
       "      <th>full_flowering_date</th>\n",
       "      <th>source_code</th>\n",
       "      <th>data_type_code</th>\n",
       "      <th>reference_name</th>\n",
       "      <th>date_time</th>\n",
       "      <th>month</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>11</th>\n",
       "      <td>812</td>\n",
       "      <td>92.0</td>\n",
       "      <td>401</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-01</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>14</th>\n",
       "      <td>815</td>\n",
       "      <td>105.0</td>\n",
       "      <td>415</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-15</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>30</th>\n",
       "      <td>831</td>\n",
       "      <td>96.0</td>\n",
       "      <td>406</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-06</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>50</th>\n",
       "      <td>851</td>\n",
       "      <td>108.0</td>\n",
       "      <td>418</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>MONTOKUTENNO-JITSUROKU</td>\n",
       "      <td>1900-04-18</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>52</th>\n",
       "      <td>853</td>\n",
       "      <td>104.0</td>\n",
       "      <td>414</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>MONTOKUTENNO-JITSUROKU</td>\n",
       "      <td>1900-04-14</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>...</th>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1210</th>\n",
       "      <td>2011</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-09</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1211</th>\n",
       "      <td>2012</td>\n",
       "      <td>101.0</td>\n",
       "      <td>410</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-10</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1212</th>\n",
       "      <td>2013</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-03</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1213</th>\n",
       "      <td>2014</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-04</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1214</th>\n",
       "      <td>2015</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-03</td>\n",
       "      <td>April</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "<p>827 rows × 8 columns</p>\n",
       "</div>"
      ],
      "text/plain": [
       "        ad  full_flowering_date_doy full_flowering_date  source_code  \\\n",
       "11     812                     92.0                 401          1.0   \n",
       "14     815                    105.0                 415          1.0   \n",
       "30     831                     96.0                 406          1.0   \n",
       "50     851                    108.0                 418          1.0   \n",
       "52     853                    104.0                 414          1.0   \n",
       "...    ...                      ...                 ...          ...   \n",
       "1210  2011                     99.0                 409          8.0   \n",
       "1211  2012                    101.0                 410          8.0   \n",
       "1212  2013                     93.0                 403          8.0   \n",
       "1213  2014                     94.0                 404          8.0   \n",
       "1214  2015                     93.0                 403          8.0   \n",
       "\n",
       "      data_type_code          reference_name  date_time  month  \n",
       "11               2.0              NIHON-KOKI 1900-04-01  April  \n",
       "14               2.0              NIHON-KOKI 1900-04-15  April  \n",
       "30               2.0              NIHON-KOKI 1900-04-06  April  \n",
       "50               2.0  MONTOKUTENNO-JITSUROKU 1900-04-18  April  \n",
       "52               2.0  MONTOKUTENNO-JITSUROKU 1900-04-14  April  \n",
       "...              ...                     ...        ...    ...  \n",
       "1210             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-09  April  \n",
       "1211             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-10  April  \n",
       "1212             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-03  April  \n",
       "1213             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-04  April  \n",
       "1214             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-03  April  \n",
       "\n",
       "[827 rows x 8 columns]"
      ]
     },
     "execution_count": 22,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 17. Using your new column, how many blossomings happened in each month?"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 23,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "April    814\n",
       "March     10\n",
       "May        3\n",
       "Name: month, dtype: int64"
      ]
     },
     "execution_count": 23,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df.month.value_counts()"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 18. Make a bar graph of how many blossomings happened in each month."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 24,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "<AxesSubplot:>"
      ]
     },
     "execution_count": 24,
     "metadata": {},
     "output_type": "execute_result"
    },
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "df.month.value_counts().plot(kind=\"bar\")"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 19. Adding a day-of-month column\n",
    "\n",
    "Now we're going to add a new column called `day_of_month.`\n",
    "\n",
    "*Tip: If you didn't drop the rows missing full-flowering dates earlier, it will yell at you about missing data. Go back up and fix Number 6!*"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 25,
   "metadata": {},
   "outputs": [],
   "source": [
    "df['day_of_month'] = df.date_time.dt.day"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 26,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>ad</th>\n",
       "      <th>full_flowering_date_doy</th>\n",
       "      <th>full_flowering_date</th>\n",
       "      <th>source_code</th>\n",
       "      <th>data_type_code</th>\n",
       "      <th>reference_name</th>\n",
       "      <th>date_time</th>\n",
       "      <th>month</th>\n",
       "      <th>day_of_month</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>11</th>\n",
       "      <td>812</td>\n",
       "      <td>92.0</td>\n",
       "      <td>401</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-01</td>\n",
       "      <td>April</td>\n",
       "      <td>1</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>14</th>\n",
       "      <td>815</td>\n",
       "      <td>105.0</td>\n",
       "      <td>415</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-15</td>\n",
       "      <td>April</td>\n",
       "      <td>15</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>30</th>\n",
       "      <td>831</td>\n",
       "      <td>96.0</td>\n",
       "      <td>406</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-06</td>\n",
       "      <td>April</td>\n",
       "      <td>6</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>50</th>\n",
       "      <td>851</td>\n",
       "      <td>108.0</td>\n",
       "      <td>418</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>MONTOKUTENNO-JITSUROKU</td>\n",
       "      <td>1900-04-18</td>\n",
       "      <td>April</td>\n",
       "      <td>18</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>52</th>\n",
       "      <td>853</td>\n",
       "      <td>104.0</td>\n",
       "      <td>414</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>MONTOKUTENNO-JITSUROKU</td>\n",
       "      <td>1900-04-14</td>\n",
       "      <td>April</td>\n",
       "      <td>14</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>...</th>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1210</th>\n",
       "      <td>2011</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-09</td>\n",
       "      <td>April</td>\n",
       "      <td>9</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1211</th>\n",
       "      <td>2012</td>\n",
       "      <td>101.0</td>\n",
       "      <td>410</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-10</td>\n",
       "      <td>April</td>\n",
       "      <td>10</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1212</th>\n",
       "      <td>2013</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-03</td>\n",
       "      <td>April</td>\n",
       "      <td>3</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1213</th>\n",
       "      <td>2014</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-04</td>\n",
       "      <td>April</td>\n",
       "      <td>4</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1214</th>\n",
       "      <td>2015</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-03</td>\n",
       "      <td>April</td>\n",
       "      <td>3</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "<p>827 rows × 9 columns</p>\n",
       "</div>"
      ],
      "text/plain": [
       "        ad  full_flowering_date_doy full_flowering_date  source_code  \\\n",
       "11     812                     92.0                 401          1.0   \n",
       "14     815                    105.0                 415          1.0   \n",
       "30     831                     96.0                 406          1.0   \n",
       "50     851                    108.0                 418          1.0   \n",
       "52     853                    104.0                 414          1.0   \n",
       "...    ...                      ...                 ...          ...   \n",
       "1210  2011                     99.0                 409          8.0   \n",
       "1211  2012                    101.0                 410          8.0   \n",
       "1212  2013                     93.0                 403          8.0   \n",
       "1213  2014                     94.0                 404          8.0   \n",
       "1214  2015                     93.0                 403          8.0   \n",
       "\n",
       "      data_type_code          reference_name  date_time  month  day_of_month  \n",
       "11               2.0              NIHON-KOKI 1900-04-01  April             1  \n",
       "14               2.0              NIHON-KOKI 1900-04-15  April            15  \n",
       "30               2.0              NIHON-KOKI 1900-04-06  April             6  \n",
       "50               2.0  MONTOKUTENNO-JITSUROKU 1900-04-18  April            18  \n",
       "52               2.0  MONTOKUTENNO-JITSUROKU 1900-04-14  April            14  \n",
       "...              ...                     ...        ...    ...           ...  \n",
       "1210             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-09  April             9  \n",
       "1211             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-10  April            10  \n",
       "1212             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-03  April             3  \n",
       "1213             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-04  April             4  \n",
       "1214             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-03  April             3  \n",
       "\n",
       "[827 rows x 9 columns]"
      ]
     },
     "execution_count": 26,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 20. Adding a date column\n",
    "\n",
    "If you don't have a nice-looking date column yet, take the `'month'` and `'day_of_month'` columns and combine them in order to create a new column called `'date'`. By \"nice looking,\" I mean it should say something like `April 11`.\n",
    "\n",
    "* Instead of using the two existing columns, you could learn to use `.dt.strftime` as mentioned above."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 33,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>ad</th>\n",
       "      <th>full_flowering_date_doy</th>\n",
       "      <th>full_flowering_date</th>\n",
       "      <th>source_code</th>\n",
       "      <th>data_type_code</th>\n",
       "      <th>reference_name</th>\n",
       "      <th>date_time</th>\n",
       "      <th>month</th>\n",
       "      <th>day_of_month</th>\n",
       "      <th>date</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>11</th>\n",
       "      <td>812</td>\n",
       "      <td>92.0</td>\n",
       "      <td>401</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-01</td>\n",
       "      <td>April</td>\n",
       "      <td>1</td>\n",
       "      <td>Apr 01</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>14</th>\n",
       "      <td>815</td>\n",
       "      <td>105.0</td>\n",
       "      <td>415</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-15</td>\n",
       "      <td>April</td>\n",
       "      <td>15</td>\n",
       "      <td>Apr 15</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>30</th>\n",
       "      <td>831</td>\n",
       "      <td>96.0</td>\n",
       "      <td>406</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>NIHON-KOKI</td>\n",
       "      <td>1900-04-06</td>\n",
       "      <td>April</td>\n",
       "      <td>6</td>\n",
       "      <td>Apr 06</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>50</th>\n",
       "      <td>851</td>\n",
       "      <td>108.0</td>\n",
       "      <td>418</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>MONTOKUTENNO-JITSUROKU</td>\n",
       "      <td>1900-04-18</td>\n",
       "      <td>April</td>\n",
       "      <td>18</td>\n",
       "      <td>Apr 18</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>52</th>\n",
       "      <td>853</td>\n",
       "      <td>104.0</td>\n",
       "      <td>414</td>\n",
       "      <td>1.0</td>\n",
       "      <td>2.0</td>\n",
       "      <td>MONTOKUTENNO-JITSUROKU</td>\n",
       "      <td>1900-04-14</td>\n",
       "      <td>April</td>\n",
       "      <td>14</td>\n",
       "      <td>Apr 14</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>...</th>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "      <td>...</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1210</th>\n",
       "      <td>2011</td>\n",
       "      <td>99.0</td>\n",
       "      <td>409</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-09</td>\n",
       "      <td>April</td>\n",
       "      <td>9</td>\n",
       "      <td>Apr 09</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1211</th>\n",
       "      <td>2012</td>\n",
       "      <td>101.0</td>\n",
       "      <td>410</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-10</td>\n",
       "      <td>April</td>\n",
       "      <td>10</td>\n",
       "      <td>Apr 10</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1212</th>\n",
       "      <td>2013</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-03</td>\n",
       "      <td>April</td>\n",
       "      <td>3</td>\n",
       "      <td>Apr 03</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1213</th>\n",
       "      <td>2014</td>\n",
       "      <td>94.0</td>\n",
       "      <td>404</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-04</td>\n",
       "      <td>April</td>\n",
       "      <td>4</td>\n",
       "      <td>Apr 04</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1214</th>\n",
       "      <td>2015</td>\n",
       "      <td>93.0</td>\n",
       "      <td>403</td>\n",
       "      <td>8.0</td>\n",
       "      <td>0.0</td>\n",
       "      <td>NEWS-PAPER(ARASHIYAMA)</td>\n",
       "      <td>1900-04-03</td>\n",
       "      <td>April</td>\n",
       "      <td>3</td>\n",
       "      <td>Apr 03</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "<p>827 rows × 10 columns</p>\n",
       "</div>"
      ],
      "text/plain": [
       "        ad  full_flowering_date_doy full_flowering_date  source_code  \\\n",
       "11     812                     92.0                 401          1.0   \n",
       "14     815                    105.0                 415          1.0   \n",
       "30     831                     96.0                 406          1.0   \n",
       "50     851                    108.0                 418          1.0   \n",
       "52     853                    104.0                 414          1.0   \n",
       "...    ...                      ...                 ...          ...   \n",
       "1210  2011                     99.0                 409          8.0   \n",
       "1211  2012                    101.0                 410          8.0   \n",
       "1212  2013                     93.0                 403          8.0   \n",
       "1213  2014                     94.0                 404          8.0   \n",
       "1214  2015                     93.0                 403          8.0   \n",
       "\n",
       "      data_type_code          reference_name  date_time  month  day_of_month  \\\n",
       "11               2.0              NIHON-KOKI 1900-04-01  April             1   \n",
       "14               2.0              NIHON-KOKI 1900-04-15  April            15   \n",
       "30               2.0              NIHON-KOKI 1900-04-06  April             6   \n",
       "50               2.0  MONTOKUTENNO-JITSUROKU 1900-04-18  April            18   \n",
       "52               2.0  MONTOKUTENNO-JITSUROKU 1900-04-14  April            14   \n",
       "...              ...                     ...        ...    ...           ...   \n",
       "1210             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-09  April             9   \n",
       "1211             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-10  April            10   \n",
       "1212             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-03  April             3   \n",
       "1213             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-04  April             4   \n",
       "1214             0.0  NEWS-PAPER(ARASHIYAMA) 1900-04-03  April             3   \n",
       "\n",
       "        date  \n",
       "11    Apr 01  \n",
       "14    Apr 15  \n",
       "30    Apr 06  \n",
       "50    Apr 18  \n",
       "52    Apr 14  \n",
       "...      ...  \n",
       "1210  Apr 09  \n",
       "1211  Apr 10  \n",
       "1212  Apr 03  \n",
       "1213  Apr 04  \n",
       "1214  Apr 03  \n",
       "\n",
       "[827 rows x 10 columns]"
      ]
     },
     "execution_count": 33,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "df['date'] = df.date_time.dt.strftime(\"%b %d\")\n",
    "df"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## 21. What day of the week do cherry blossoms like to blossom on?\n",
    "\n",
    "Do they get the weekends off?"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 34,
   "metadata": {},
   "outputs": [
    {
     "ename": "AttributeError",
     "evalue": "'Series' object has no attribute 'values_counts'",
     "output_type": "error",
     "traceback": [
      "\u001b[0;31m---------------------------------------------------------------------------\u001b[0m",
      "\u001b[0;31mAttributeError\u001b[0m                            Traceback (most recent call last)",
      "Input \u001b[0;32mIn [34]\u001b[0m, in \u001b[0;36m<cell line: 1>\u001b[0;34m()\u001b[0m\n\u001b[0;32m----> 1\u001b[0m \u001b[43mdf\u001b[49m\u001b[38;5;241;43m.\u001b[39;49m\u001b[43mdate\u001b[49m\u001b[38;5;241;43m.\u001b[39;49m\u001b[43mvalues_counts\u001b[49m()\n",
      "File \u001b[0;32m~/.pyenv/versions/3.10.3/lib/python3.10/site-packages/pandas/core/generic.py:5575\u001b[0m, in \u001b[0;36mNDFrame.__getattr__\u001b[0;34m(self, name)\u001b[0m\n\u001b[1;32m   5568\u001b[0m \u001b[38;5;28;01mif\u001b[39;00m (\n\u001b[1;32m   5569\u001b[0m     name \u001b[38;5;129;01mnot\u001b[39;00m \u001b[38;5;129;01min\u001b[39;00m \u001b[38;5;28mself\u001b[39m\u001b[38;5;241m.\u001b[39m_internal_names_set\n\u001b[1;32m   5570\u001b[0m     \u001b[38;5;129;01mand\u001b[39;00m name \u001b[38;5;129;01mnot\u001b[39;00m \u001b[38;5;129;01min\u001b[39;00m \u001b[38;5;28mself\u001b[39m\u001b[38;5;241m.\u001b[39m_metadata\n\u001b[1;32m   5571\u001b[0m     \u001b[38;5;129;01mand\u001b[39;00m name \u001b[38;5;129;01mnot\u001b[39;00m \u001b[38;5;129;01min\u001b[39;00m \u001b[38;5;28mself\u001b[39m\u001b[38;5;241m.\u001b[39m_accessors\n\u001b[1;32m   5572\u001b[0m     \u001b[38;5;129;01mand\u001b[39;00m \u001b[38;5;28mself\u001b[39m\u001b[38;5;241m.\u001b[39m_info_axis\u001b[38;5;241m.\u001b[39m_can_hold_identifiers_and_holds_name(name)\n\u001b[1;32m   5573\u001b[0m ):\n\u001b[1;32m   5574\u001b[0m     \u001b[38;5;28;01mreturn\u001b[39;00m \u001b[38;5;28mself\u001b[39m[name]\n\u001b[0;32m-> 5575\u001b[0m \u001b[38;5;28;01mreturn\u001b[39;00m \u001b[38;5;28;43mobject\u001b[39;49m\u001b[38;5;241;43m.\u001b[39;49m\u001b[38;5;21;43m__getattribute__\u001b[39;49m\u001b[43m(\u001b[49m\u001b[38;5;28;43mself\u001b[39;49m\u001b[43m,\u001b[49m\u001b[43m \u001b[49m\u001b[43mname\u001b[49m\u001b[43m)\u001b[49m\n",
      "\u001b[0;31mAttributeError\u001b[0m: 'Series' object has no attribute 'values_counts'"
     ]
    }
   ],
   "source": [
    "df.date.values_counts()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# YOU ARE DONE.\n",
    "\n",
    "And **incredible.**"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.10.3"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 2
}
